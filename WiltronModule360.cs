/*************************************************************************
 * 
 *               File     : WiltronModule360.cs
 *               Class    : Wiltron360
 *               Library  : Metas.Instr.Driver.Vna
 *               Version  : 1.0.0.0
 *               Author   : Andre Hoehne
 *               Created  : 10.02.2023
 *               Modified : --.--.----
 * 
 ************************************************************************/

using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text;
using Ivi.Visa;
using Metas.Instr.VisaExtensions;
using System.Text;


using Metas.UncLib.Core;
using Metas.Vna.Data;

using DoubleHelper = Metas.UncLib.Core.Special.DoubleHelper;
using Metas.Vna.Data.FileIO;

namespace Metas.Instr.Driver.Vna
{
    internal class Wiltron360 : IVna, ISrqDevice
    {
        /// <summary>
        /// Visa session.
        /// </summary>
        public SrqMessageBasedSession visa;

        private VnaSetUpMode setupmode;
        internal VnaParameter[] parameters;

        private string stateFile = "";
        internal int nTestPorts = 2;
        private bool disp4p;
        bool TriggerHoldActive = false;
        #region Constructor

        public Wiltron360()
        {
            this.parameters = new VnaParameter[0];
        }
        public Wiltron360(string resourceName, string idName = null)
        {
            Open(resourceName, idName);
        }
        #endregion

        #region IDevice Members

        /// <summary>
        /// Opens a VNA session to the specified resource.
        /// </summary>
        /// <param name="resourceName">String that describes a unique VISA resource.</param>
        /// <param name="idName">String that describes the instrument identification.</param>
        public void Open(string resourceName, string idName = null)
        {
        //    throw new NotImplementedException();
            this.setupmode = VnaSetUpMode.unknown;
            this.parameters = new VnaParameter[0];
        //    this.parameterMatrix = null;
            this.visa = new SrqMessageBasedSession(resourceName);
            string idn = Identification;

            this.visa.SrqTimeout = 10000;
            // Enable SQRs / Enable any unmaked service request
            this.visa.Write("SQ1");
            this.visa.Timeout = 10000;
            this.visa.SrqMask = "SQ1";
            this.visa.SrqTimeout = 10000;

        }
        public void Close()
        {
            this.visa.Write("RST RTL");
        }

        #endregion

        #region IVna Members

        private string _rootPath = "";

        /// <summary>
        /// Root Path
        /// </summary>
        public string RootPath
        {
            get { return _rootPath; }
            set { _rootPath = value; }
        }

        /// <summary>
        /// Number of Test Ports
        /// </summary>
        public int NTestPorts
        {
            get
            {
                return nTestPorts;
            }
        }
        //![0] Sweepmode default on Linear Frequency mode
        public VnaSweepMode SweepMode
        {
            get
            {
                double[] list = ConvertStringList2DoubleList( this.visa.Query("OFV").Split('\n'));
                int CntPoint = list.Length-1;
                double sum = 0;
                foreach(double x in list)
                {
                    if (x == Double.NaN) continue;
                    sum += x;
                }
                double avg = sum / CntPoint;

                if (avg == list[0])
                {
                    return VnaSweepMode.CWTime;
                }
                else
                {
                    return VnaSweepMode.LinearFrequency;
                }
            }
            set
            {
                switch (value)
                {
                    case VnaSweepMode.LinearFrequency:
                        this.FrequencyStart = 40e6;
                        this.FrequencyStop = 40e9;
                        break;
                    case VnaSweepMode.CWTime:
                        this.visa.Write("CWF 1GHZ");
                        break;
                    default:
                        throw new System.Exception("VNA Sweep Mode not supported.");
                }

            }
        }

        /// <summary>
        /// Sweep Time / s
        /// </summary>
        public double SweepTime
        { 
            get
            {
                return double.NaN;
            }
            set 
            {

            }

        }

        /// <summary>
        /// Dwell Time / s
        /// </summary>
        public double DwellTime
        {
            get
            {
                return double.NaN;
            }
            set
            {
                
            }
        }

        /// <summary>
        /// IF Average Factor
        /// </summary>
        public int IFAverageFactor
        {
            get
            {
                return 0;
            }
            set
            {
                if (value <= 1)
                {
                    this.visa.Write("AOF");
                }
                else
                {
                    this.visa.Write("AVG " + value.ToString() + "XX1");
                }
            }
        }

        /// <summary>
        /// Start Frequency / Hz
        /// </summary>
        public double FrequencyStart
        { get
            {
                return FrequencyList[0];
            }
            set
            {
                if (value <= 40e6)
                {
                    this.visa.Write("SRT 40 MHZ");
                }
                else
                {
                    //Convert Hz to MHz
                    double frequency = value * 1e-6;
                    this.visa.Write("SRT " + frequency.ToString() + "MHZ");
                }

            }

        }

        /// <summary>
        /// Stop Frequency / Hz
        /// </summary>
        public double FrequencyStop
        { 
            get
            {
                return FrequencyList.Last();
            }
            set
            {
                //Convert Hz to MHz
                double frequency = value * 1e-6;
                this.visa.Write("STP " + frequency.ToString() + "MHZ");
            }
                
        }

        /// <summary>
        /// Center Frequency / Hz
        /// </summary>
        public double FrequencyCenter 
        {
            get 
            {
                int len = FrequencyList.Length;
                int halflen = (int)System.Math.Floor(System.Convert.ToDouble(len,null) / 2);
                return FrequencyList[halflen];
            }
            set 
            {
            }
        }

        /// <summary>
        /// Span Frequency / Hz
        /// </summary>
        public double FrequencySpan 
        { 
            get
            {
                return FrequencyList[FrequencyList.Length-1] - FrequencyList[0];
             }
            set 
            {
            }
        }
        /// <summary>
        /// Source 1 Power / dBm
        /// </summary>
        public double Source1Power 
        {
            get
            {
                return double.NaN;
            }
            set
            {
                this.visa.SrqWrite("PWR " + value + "DBM SR1");
            }
        }

        /// <summary>
        /// Source 2 Power / dBm
        /// </summary>
        public double Source2Power
        {
            get
            {
                return double.NaN;
            }
            set 
            {
                this.visa.SrqWrite("PW2 " + value.ToString() + "DBM SR1");
            }
        }
        /// <summary>
        /// Port 1 Attenuator / dB
        /// </summary>
        public double Port1Attenuator {
            get
            {
                return double.NaN;
            }
            set
            {
                this.visa.SrqWrite("SA1 " + value.ToString() + "DBL SR1");
            }
        }
        /// <summary>
        /// Port 2 Attenuator / dB
        /// </summary>
        public double Port2Attenuator
        {
            get
            {
                return double.NaN;
            }
            set
            {
                this.visa.SrqWrite("SA2 " + value.ToString() + "DBL SR1");
            }
        }
        public double FrequencyCW
        {
            get 
            {
                VnaSweepMode sweepMode = SweepMode;
                if(sweepMode == VnaSweepMode.CWTime)  return VnaHelper.String2Double(this.visa.Query("OFV").Trim().Split('\n')[0]);
                else return 0;
                
            }
            set
            {
                SweepMode = VnaSweepMode.CWTime;
                double Frequency = value *1e-6;
                this.visa.Write("CWF "+ Frequency.ToString()+" MHZ");
            }
        }

        public int SweepPoints { 
            get
            {
                
                return this.visa.Query("OFV").Trim().Split('\n').Length;
            }
            set 
            {
                VnaSweepMode sweepMode = SweepMode;
                if (sweepMode == VnaSweepMode.CWTime)
                {
                    this.visa.Write("CWP " + value.ToString()+" XX1");
                }

            } 
        }
        public double IFBandwidth {
            get
            {
                return double.NaN;
            }
            set
            {
                if (value <= 10) this.visa.Write("IFM");
                else if (value == 100) this.visa.Write("IFN");
                else this.visa.Write("IFR");
            }

        }
        public double[] FrequencyList
        {
            get
            {
                string[] strfrequencylist = this.visa.Query("OFV").Trim().Split('\n');
                List<double> frequnecy=new List<double>();
                foreach(string strfrequency in strfrequencylist) {                
                    frequnecy.Add(VnaHelper.String2Double(strfrequency));  
                }
                return frequnecy.ToArray();
            }
            set
            {

            }
        }

        #endregion


        public VnaSettings Settings => throw new NotImplementedException();
        public double[,] SegmentTable {
            get { return new double[0,0]; }
            set { } }
        public double Source1PowerSlope { 
            get { return 0.0; }
            set { }
        }
        public double Port1Extension
        {
            get { return 0.0; }
            set { }
        }
        public double Source2PowerSlope
        {
            get { return 0.0; }
            set { }
        }
        public double Port2Extension
        {
            get { return 0.0; }
            set { }
        }
        /// <summary>
        /// Source Port Settings
        /// </summary>
        public VnaSourcePortSettings[] SourcePorts
        {
            get
            {
                return VnaHelper.GetVnaSourcePortSettingsFromObsoleteProperties(this);
            }
            set
            {
                VnaHelper.SetVnaSourcePortSettingsToObsoleteProperties(this, value);
            }
        }
        /// <summary>
        /// Z0 / Ohm
        /// </summary>
        public double Z0 {
            get { return 50.0; }
            set { }
        }
        public bool OutputState
        {
            get { return true; }
            set
            {
                if (value)
                {
                    this.visa.Write("RH1");
                }
                else
                {
                    this.visa.Write("RH0");
                }
            }
        }

        private VnaParameter[,] parameterMatrix = null;
        /// <summary>
        /// VNA parameter matrix
        /// </summary>
        public VnaParameter[,] ParameterMatrix
        {
            get
            {
                return parameterMatrix;
            }
            set
            {
                parameterMatrix = null;
                SetUp(VnaHelper.GetVnaSetUpMode(value));
                parameterMatrix = value;
            }
        }

        public VnaSetUpMode SetUpMode => throw new NotImplementedException();

        public VnaParameter[] Parameters => throw new NotImplementedException();

        public string Identification
        {
            get
            {
                return this.visa.Query("OID");
            }
        }

        public string SrqIdentification
        {
            get
            {
                return this.visa.SrqQuery("OID SR1");
            }
        }

    public VnaData<Number> GetData(VnaFormat format)
        {
            throw new NotImplementedException();
        }

        public byte[] GetState()
        {
            throw new NotImplementedException();
        }

        public bool GetState(string pathState)
        {
            throw new NotImplementedException();
        }

        public void Measure(string pathRes)
        {
            throw new NotImplementedException();
        }

        public void Preset()
        {
            throw new NotImplementedException();
        }

        public void SetSingleFrequency(double freq)
        {
            throw new NotImplementedException();
        }

        public void SetSingleFrequency(double[] freq, int index)
        {
            throw new NotImplementedException();
        }

        public void SetState(byte[] state)
        {
            throw new NotImplementedException();
        }

        public bool SetState(string pathState)
        {
            throw new NotImplementedException();
        }

        public void SetUp(VnaSetUpMode mode)
        {
            switch (mode)
            {
                case VnaSetUpMode.unknown:
                    throw new System.Exception("VNA SetUp Mode unknown not supported.");
                case VnaSetUpMode.S11:
                    parameters = new VnaParameter[] { VnaParameter.SParameter(1, 1) };
                    this.disp4p = false;
                    this.visa.Write("DSP S11");
                    break;
                case VnaSetUpMode.S21:
                    parameters = new VnaParameter[] { VnaParameter.SParameter(2, 1) };
                    this.disp4p = false;
                    this.visa.Write("DSP S21");
                    break;
                case VnaSetUpMode.S12:
                    parameters = new VnaParameter[] { VnaParameter.SParameter(1, 2) };
                    this.disp4p = false;
                    this.visa.Write("DSP S12");
                    break;
                case VnaSetUpMode.S22:
                    parameters = new VnaParameter[] { VnaParameter.SParameter(2, 2) };
                    this.disp4p = false;
                    this.visa.Write("DSP S22");
                    break;
                case VnaSetUpMode.Sxx:
                    parameters = new VnaParameter[] { VnaParameter.SParameter(1, 1),
                                                      VnaParameter.SParameter(2, 1),
                                                      VnaParameter.SParameter(1, 2),
                                                      VnaParameter.SParameter(2, 2) };
                    this.disp4p = true;
                    this.visa.Write("D14 CH1 S11 CH2 S12 CH3 S21 CH4 S22");
                    break;
            }
        }

        /// <summary>
        /// Set up VNA.
        /// </summary>
        /// <param name="mode">VNA mode</param>
        public void SetUp(string mode)
        {
            SetUp((VnaSetUpMode)System.Enum.Parse(typeof(VnaSetUpMode), mode));
        }

        public void TriggerCont()
        {
            this.visa.Write("CTN");
        }

        public void TriggerHold()
        {
            this.visa.Write("HLD");
        }

        public void TriggerSingle(BackgroundWorker worker = null, DoWorkEventArgs e = null)
        {

            this.visa.Write("TRS");
        }

        
        int srq_timeout;

            /// <summary>
            /// Trigger device.
            /// </summary>
        public void TriggerSingleStart()
        {
             srq_timeout = this.visa.SrqTimeout;
             this.visa.SrqMask = "SQ1";
             this.visa.SrqTimeout = 7200000;
             this.visa.SrqBegin();
             this.visa.Write("TRS");
             
        }
        

        public void TriggerSingleWait(BackgroundWorker worker = null, DoWorkEventArgs e = null)
        {
            throw new NotImplementedException();
        }
    }
}
