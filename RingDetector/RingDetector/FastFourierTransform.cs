using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using MathNet.Numerics.Transformations;
using NICEAutomation.AutomationCommon;
using NICEAutomation.AutomationCommon.Output;

namespace CallValidation
{
    public class FastFourierTransform
    {
        private IOutputStreamer myOutput;

        public FastFourierTransform()
        {
            myOutput = OutputStreamer.GetInstance().GetCurrentStreamer();
        }


        [ComVisible(true)]
        public List<string> CountSentences(IOutputStreamer _objTestLog, string filePath) //MP
        {
            int sampleRate = 0; //Hz
            double[] allSamples = GetSamplesForMonoWav(filePath, out sampleRate);
            List<double[]> partisipantSegments = GetPartisipantSegments(allSamples, sampleRate);

            List<string> callPartisipants = DetectCallParticipants(partisipantSegments, _objTestLog, sampleRate);
            return callPartisipants;
        }

        /// <summary>
        /// Counts duration of sound segment
        /// </summary>
        /// <param name="segmentSamples"></param>
        /// <param name="sampleRate"></param>
        /// <returns></returns>
        public double GetSegmentDuration(double[] segmentSamples, int sampleRate)
        {
            double numberOfSamples = segmentSamples.Count();
            double segmentDuration = numberOfSamples / sampleRate;
            myOutput.Debug(string.Format("FastFourierTransform|GetDuration(): duration is '{0} s'", segmentDuration));

            return segmentDuration;
        }

        /// <summary>
        /// Returns array of samples for mono wav file
        /// </summary>
        /// <param name="wavePath"></param>
        /// <param name="sampleRate"></param>
        /// <returns></returns>
        public Double[] GetSamplesForMonoWav(String wavePath, out int sampleRate)
        {
            Double[] data;
            byte[] wave;
            try
            {
                System.IO.FileStream waveFile = System.IO.File.OpenRead(wavePath);
                wave = new byte[waveFile.Length];
                waveFile.Read(wave, 0, Convert.ToInt32(waveFile.Length)); //read the wave file into the wave variable;
                /***********Converting and PCM accounting***************/
                int bitsPerSample = BitConverter.ToInt16(wave, 34); //in wav file bitsPerSample info is in 34-35 bytes;
                double maxSampleValue = Math.Pow(2, bitsPerSample - 1); //-1 because samples are signed;
                int bytesForOneSample = bitsPerSample / 8;
                data = new Double[(wave.Length - 44) / bytesForOneSample]; //shifting the headers out of the PCM data;

                for (int i = 0; i < data.Length; i++)
                {
                    data[i] = (BitConverter.ToInt16(wave, i * bytesForOneSample)) / maxSampleValue;
                }

                int headerBytesToSkip = 44 / bytesForOneSample; // header length is 44 bytes;
                data = data.Skip(headerBytesToSkip).ToArray();
                /**************assigning sample rate**********************/
                sampleRate = BitConverter.ToInt32(wave, 24);    //in wav file sampleRate info is in 24-27 bytes;
                return data;
            }
            catch (Exception e)
            {
                myOutput.ReportException(e, "Failed to get samples from wav file " + wavePath);
                sampleRate = -1;
            }

            return null;
        }

        /// <summary>
        /// Detects segments with sound activity in wav and returns separated arrays of samples for each detected segment
        /// </summary>
        /// <param name="allSamples"></param>
        /// <param name="sampleRate"></param>
        /// <param name="minSilenceDurationBetweenSegments">minimum silence duration between segments, sec. If less - will be assumed as one divided segment</param>
        /// <returns></returns>
        public List<double[]> GetPartisipantSegments(double[] allSamples, int sampleRate, double minSilenceDurationBetweenSegments = 0.5)
        {
            List<double[]> partisipantSegments = new List<double[]>();
            double[] tempPartisipantSegment;
            double minVoiceActivityPeriod = 0.05; //Will be calculated Root Mean Square for each 0.05 seconds;
            int samplesCountForRMS = Convert.ToInt32(sampleRate * minVoiceActivityPeriod);
            double[] samplesRange = new double[samplesCountForRMS];
            int startPosition = 0;
            int numberOfDetectedVoiceActivity = 0;
            int stopPosition;
            int numberOfDetectedVoiceSilence = 0;

            int maxAllowedNumberOfDetectedVoiceSilence = (int)((1 / minVoiceActivityPeriod) * (minSilenceDurationBetweenSegments));
            bool partisipantSegmentStarted = false;

            for (int i = 0; i + samplesCountForRMS < allSamples.Count(); i += samplesCountForRMS)
            {
                samplesRange = allSamples.Skip(i).Take(samplesCountForRMS).ToArray();
                if (DetectVoiceActivity(samplesRange))
                {
                    numberOfDetectedVoiceActivity++;
                    startPosition = !partisipantSegmentStarted ? i : startPosition;
                    partisipantSegmentStarted = true;
                    numberOfDetectedVoiceSilence = 0;
                }
                else if (partisipantSegmentStarted)
                {
                    numberOfDetectedVoiceSilence++;
                    if (numberOfDetectedVoiceActivity <= 1) //minVoiceActivityPeriod * 1 = 0.1 seconds - if audio activity less - assume it as noise;
                    {
                        partisipantSegmentStarted = false;
                        numberOfDetectedVoiceActivity = 0;
                        continue;
                    }
                    if (numberOfDetectedVoiceSilence > maxAllowedNumberOfDetectedVoiceSilence)
                    {
                        stopPosition = i - ((maxAllowedNumberOfDetectedVoiceSilence) * samplesCountForRMS);
                        tempPartisipantSegment = allSamples.Skip(startPosition).Take(stopPosition - startPosition).ToArray();
                        partisipantSegments.Add(tempPartisipantSegment);
                        partisipantSegmentStarted = false;
                        numberOfDetectedVoiceActivity = 0;
                    }
                }
            }

            foreach (var segment in partisipantSegments)
            {
                myOutput.Debug(string.Format("Was detected segment with duration '{0} s'", segment.Count() / sampleRate));
            }
            myOutput.Debug(string.Format("FastFourierTransform|GetPartisipantSegments() returned '{0}' segments", partisipantSegments.Count));
            return partisipantSegments;
        }

        /// <summary>
        /// Returns segments with activity
        /// </summary>
        /// <param name="samplesRange"></param>
        /// <param name="signalLowLevelLimit">low limit for detecting signal. Default value is 1% from max possible</param>
        /// <returns></returns>
        private bool DetectVoiceActivity(double[] samplesRange, double signalLowLevelLimit = 0.01)
        {
            double squaredSum = 0;
            double rootMeanSquare = 0;

            samplesRange.ToList().ForEach(sample => squaredSum += sample * sample);
            rootMeanSquare = Math.Sqrt(squaredSum / samplesRange.Count());

            return rootMeanSquare > signalLowLevelLimit;
        }

        private List<string> DetectCallParticipants(List<double[]> partisipantSegments, IOutputStreamer _objTestLog, int sampleRate)
        {
            List<string> participants = new List<string>();
            double duration;
            int NumberCount = 0;

            foreach (double[] segment in partisipantSegments)
            {
                duration = GetSegmentDuration(segment, sampleRate);
                double tollerance = 0.5; //0.5 seconds
                string part = String.Empty;

                if (duration <= 2)
                {
                    _objTestLog.WriteLog(ClsEnum.ReportTypes.DEBUG, part + "Ring tone detected. Duration : " + duration, false);
                    continue;
                }
                if (duration <= 3 + tollerance)
                {
                    participants.Add("A1");
                    part = "A1";
                    _objTestLog.WriteLog(ClsEnum.ReportTypes.DEBUG, part + " Duration : " + duration, false);
                    NumberCount++;
                    continue;
                }
                if (duration <= (4 + tollerance))
                {
                    participants.Add("A1_SOD");
                    part = "A1_SOD";
                    _objTestLog.WriteLog(ClsEnum.ReportTypes.DEBUG, part + " Duration : " + duration, false);
                    NumberCount++;
                    continue;
                }
                if (duration <= (5 + tollerance))
                {
                    participants.Add("A2");
                    part = "A2";
                    _objTestLog.WriteLog(ClsEnum.ReportTypes.DEBUG, part + " Duration : " + duration, false);
                    NumberCount++;
                    continue;
                }
                if (duration <= (6 + tollerance))
                {
                    participants.Add("A2_SOD");
                    part = "A2_SOD";
                    _objTestLog.WriteLog(ClsEnum.ReportTypes.DEBUG, part + " Duration : " + duration, false);
                    NumberCount++;
                    continue;
                }
                if (duration <= (7 + tollerance))
                {
                    participants.Add("C1");
                    part = "C1";
                    _objTestLog.WriteLog(ClsEnum.ReportTypes.DEBUG, part + " Duration : " + duration, false);
                    NumberCount++;
                    continue;
                }

                if (duration <= (8 + tollerance))
                {
                    participants.Add("C1_SOD");
                    part = "C1_SOD";
                    _objTestLog.WriteLog(ClsEnum.ReportTypes.DEBUG, part + " Duration : " + duration, false);
                    NumberCount++;
                    continue;
                }
                if (duration <= (9 + tollerance))
                {
                    participants.Add("C2");
                    part = "C2";
                    _objTestLog.WriteLog(ClsEnum.ReportTypes.DEBUG, part + " Duration : " + duration, false);
                    NumberCount++;
                    continue;
                }
                if (duration <= (10 + tollerance))
                {
                    participants.Add("C2_SOD");
                    part = "C2_SOD";
                    _objTestLog.WriteLog(ClsEnum.ReportTypes.DEBUG, part + " Duration : " + duration, false);
                    NumberCount++;
                    continue;
                }
                if (duration <= (11 + tollerance))
                {
                    participants.Add("SV1");
                    part = "SV1";
                    _objTestLog.WriteLog(ClsEnum.ReportTypes.DEBUG, part + " Duration : " + duration, false);
                    NumberCount++;
                    continue;
                }
                if (duration <= (12 + tollerance))
                {
                    participants.Add("SV1_SOD");
                    part = "SV1_SOD";
                    _objTestLog.WriteLog(ClsEnum.ReportTypes.DEBUG, part + " Duration : " + duration, false);
                    NumberCount++;
                    continue;
                }
                if (duration <= (13 + tollerance))
                {
                    participants.Add("SV2");
                    part = "SV2";
                    _objTestLog.WriteLog(ClsEnum.ReportTypes.DEBUG, part + " Duration : " + duration, false);
                    NumberCount++;
                    continue;
                }
                if (duration <= (14 + tollerance))
                {
                    participants.Add("SV2_SOD");
                    part = "SV2_SOD";
                    _objTestLog.WriteLog(ClsEnum.ReportTypes.DEBUG, part + " Duration : " + duration, false);
                    NumberCount++;
                    continue;
                }
                if (duration <= (15 + tollerance))
                {
                    participants.Add("A3");
                    part = "A3";
                    _objTestLog.WriteLog(ClsEnum.ReportTypes.DEBUG, part + " Duration : " + duration, false);
                    NumberCount++;
                    continue;
                }
                if (duration <= (16 + tollerance))
                {
                    participants.Add("A3_SOD");
                    part = "A3_SOD";
                    _objTestLog.WriteLog(ClsEnum.ReportTypes.DEBUG, part + " Duration : " + duration, false);
                    NumberCount++;
                    continue;
                }
            }
            participants.Add(NumberCount.ToString());

            return participants;
        }
    }
}
