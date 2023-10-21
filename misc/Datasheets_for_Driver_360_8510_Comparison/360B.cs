/*************************************************************************
 * 
 *   #########   Schweizerische Eidgenossenschaft
 *   ###   ###   Confederation suisse
 *   #       #   Confederazione Svizzera
 *   ###   ###   Confederazione Svizzera
 *     #####
 *               Federal Institute of Metrology METAS
 *               Copy from:	8510C.cs 
 *				 Author	  : Michael Wollensack
 *************************************************************************
 *               File     : 360B.cs
 *               Class    : 360B
 *               Library  : Metas.Instr.Driver.Vna
 *               Version  : 0.1
 *               Author   : Daniel Bevc
 *               Created  : 09.11.2019
 *               Modified : .2019
 * 
 ************************************************************************/

using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Globalization;

using Metas.Instr.VisaExtensions;
using Metas.Instr.Driver.Vna.Data;
using Metas.UncLib.Core;
using Metas.Vna.Data;

namespace Metas.Instr.Driver.Vna
{
    /// <summary>
    /// Driver for 360B vector network analyzer.
    /// </summary>
    [Guid("3c986f08-57e6-4130-b0d6-cc4d03eb8104")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComVisible(true)]
    [ComDefaultInterface(typeof(IVna))]
    public class 360B : IVna
    {
        /// <summary>
        /// Visa session.
        /// </summary>
        public SrqMessageBasedSession visa;

        private VnaSetUpMode setupmode;
        private VnaParameter[] parameters;

        private bool disp4p;

        #region Constructor

        /// <summary>
        /// Creates a Vna session.
        /// </summary>
        public 360B()
        {
            this.parameters = new VnaParameter[0];
        }

        /// <summary>
        /// Creates a Vna session to the specified resource.
        /// </summary>
        /// <param name="resourceName">String that describes a unique VISA resource.</param>
        public 360B(string resourceName)
        {
            Open(resourceName);
        }

        #endregion

        #region IVna Members

        /// <summary>
        /// Opens a Vna session to the specified resource.
        /// </summary>
        /// <param name="resourceName">String that describes a unique VISA resource.</param>
        public void Open(string resourceName) /// OPEN VISA
        {
            this.setupmode = VnaSetUpMode.unknown;
            this.parameters = new VnaParameter[0];
            this.parameterMatrix = null;
            this.visa = new SrqMessageBasedSession(resourceName);
            string idn = Identification;
            if (idn.IndexOf("360B") < 0)
            {
                throw new System.Exception("Instrument identification query failed.");
            }
            this.visa.Timeout = 10000;
            this.visa.SrqMask = "CSB SRQ 128"; 
            /// Clear Status Byte to 128; clear SRQ --- 
            ///;/ Send two integer ASCII values, 0 to 255 to set the service request mask.
            this.visa.SrqTimeout = 10000;
            this.visa.SrqWrite("TITL \".net Driver VNA - Wiltron 360B\""); 
            ///schreibt als Titel ".net Driver VNA - Wiltron 360B"
            this.visa.SrqWrite("BACI 0;INTE 80;ENTO");
            /// BAClightIllumination off; INTEnsity 80% ; ENTryOff   
        }

        /// <summary>
        /// Closes the specified device session.
        /// </summary>
        public void Close() /// CLOSE VISA
        {
            this.TriggerCont();
            /*
			this.visa.SrqWrite("BACI 0;INTE 50;ENTO");
            */// BAClightIllumination off; INTEnsity 50% ; ENTryOff
            this.visa.Write("CSB SRQ 0 RTL"); 
            /// Clear Status Byte to 0,0; clear SRQ 
            ///;/ Send two integer ASCII values, 0 to 255 to set the service request mask./;/ 
            this.visa.Dispose();
            this.visa = null;
        }

        /// <summary>
        /// Preset the Vna.
        /// </summary>
        public void Preset()
        {
            this.setupmode = VnaSetUpMode.unknown;
            this.parameters = new VnaParameter[0];
            this.parameterMatrix = null;
            this.visa.SrqWrite("FACTPRES");
            /// Performs a FACTory PRESet (which selects frequency domain and step sweep mode)
            /*this.visa.SrqWrite("TITL \".net Driver VNA - Wiltron 360B\"");
            this.visa.SrqWrite("BACI 0;INTE 80;ENTO");
			*/
        }

        /// <summary>
        /// Set instrument state to Vna.
        /// </summary>
        /// <param name="state">Instrument state</param>
        public void SetState(byte[] state)
        {
            this.setupmode = VnaSetUpMode.unknown;
            this.parameters = new VnaParameter[0];
            this.parameterMatrix = null;
            int t = this.visa.Timeout;
            this.visa.SrqWrite("IFP");
            ///transferring front panel setup to 360 -  -writes the instrument state
            this.visa.Timeout = 30000;
            this.visa.Write(state);
            this.visa.Timeout = t;
            /*this.visa.SrqWrite("TITL \".net Driver VNA - Wiltron 360B\"");
            this.visa.SrqWrite("BACI 0;INTE 80;ENTO");
			*/
        }

        /// <summary>
        /// Get instrument state from Vna.
        /// </summary>
        /// <returns>Instrument state</returns>
        public byte[] GetState()
        {
            this.visa.SrqWrite("OFP");
            ///transferring front panel setup to controller - reads the instrument state
            byte[] state = this.visa.ReadByteArray();
            byte[] temp = new byte[2];
            temp[0] = state[3];
            temp[1] = state[2];
            int length = System.BitConverter.ToUInt16(temp, 0) + 4;
            int lenght2 = state.Length;
            return state;
        }

        /// <summary>
        /// Set instrument state to Vna.
        /// </summary>
        /// <param name="pathState">Instrument State File Path (*.is)</param>
        /// <returns>Canceled</returns>
        public bool SetState(string pathState)
        {
            return VnaHelper.SetState(this, pathState);
        }

        /// <summary>
        /// Get instrument state from Vna.
        /// </summary>
        /// <param name="pathState">Instrument State File Path (*.is)</param>
        /// <returns>Canceled</returns>
        public bool GetState(string pathState)
        {
            return VnaHelper.GetState(this, pathState);
        }

        /// <summary>
        /// Set up Vna.
        /// </summary>
        /// <param name="mode">Vna mode</param>
        public void SetUp(VnaSetUpMode mode)
        {
            switch (mode)
            {
                case VnaSetUpMode.unknown:
                    throw new System.Exception("VNA SetUp Mode unknown not supported.");
                case VnaSetUpMode.S11:
                    parameters = new VnaParameter[] { VnaParameter.SParameter(1, 1) };
                    this.disp4p = false;
                    this.visa.SrqWrite("DSP CH1 MPH"); //"SINC;S11;ENTO"
                    ///Single Channel display, Channel 1 , Mag Phase
                    break;
                case VnaSetUpMode.S21:
                    parameters = new VnaParameter[] { VnaParameter.SParameter(2, 1) };
                    this.disp4p = false;
                    this.visa.SrqWrite("DSP CH2 MPH"); //"SINC;S21;ENTO"
                    break;
                case VnaSetUpMode.S12:
                    parameters = new VnaParameter[] { VnaParameter.SParameter(1, 2) };
                    this.disp4p = false;
                    this.visa.SrqWrite("DSP CH3 MPH"); //"SINC;S12;ENTO"
                    break;
                case VnaSetUpMode.S22:
                    parameters = new VnaParameter[] { VnaParameter.SParameter(2, 2) };
                    this.disp4p = false;
                    this.visa.SrqWrite("DSP CH4 S22 MPH"); //"SINC;S22;ENTO"
                    break;
                case VnaSetUpMode.Sxx:
                    parameters = new VnaParameter[] { VnaParameter.SParameter(1, 1),
                                                      VnaParameter.SParameter(2, 1),
                                                      VnaParameter.SParameter(1, 2),
                                                      VnaParameter.SParameter(2, 2) };
                    this.disp4p = true;
                    this.visa.SrqWrite("D14 CH1 S11 MPH CH2 S12 MPH CH3 S21 MPH CH4 S22 MPH");

                    ///Select Four Parameter Split display format.
                    break;
                case VnaSetUpMode.SwitchTerms:
                    parameters = new VnaParameter[] { 
                        VnaParameter.Ratio(VnaReceiverType.a, 1, VnaReceiverType.b, 1, 2),
                        VnaParameter.Ratio(VnaReceiverType.a, 2, VnaReceiverType.b, 2, 1)
                    };
                    this.disp4p = true;
                    // Wiltron 360B if a1 is selected, no DENOMINATOR is allowed
                    this.visa.SrqWrite("D14 CH1 US2 MPH CH3 S21 MPH CH2 S12 MPH CH4 US4 MPH"); //"FOUPSPLI;USER1;S21;S12;USER4;ENTO"
                    this.visa.SrqWrite("US2 NA1 DB1 LA2 " ); //"USER1;DRIVPORT2;LOCKA2;DENOA1;NUMEB1;"
                                       //"CONV1S;PARL \"a1/b1_p2\";REDD");      "CONV1S;PARL \"a1/b1_p2\";REDD"
                    this.visa.SrqWrite("US4 NA2 DB2 LA1 " +
										"USL a2/b2"); //"USER4;DRIVPORT1;LOCKA1;DENOA2;NUMEB2;"
                                       //"CONV1S;PARL \"a2/b2_p1\";REDD"); //"CONV1S;PARL \"a2/b2_p1\";REDD"
                    
					break;
                case VnaSetUpMode.b1_b2_p1:
                    parameters = new VnaParameter[] { VnaParameter.Ratio(VnaReceiverType.b, 1, VnaReceiverType.b, 2, 1) };
                    this.disp4p = false;
                    // Wiltron 360B DENOB2 not possible
                    this.visa.SrqWrite("DSP  US2 NB1 DB2 LA1" + //"SINC;USER1;DRIVPORT1;LOCKA1;NUMEB2;DENOB1;"
                                       "USL b1/b2"); // "CONV1S;PARL \"b1/b2_p1\";REDD;ENTO"
                    break;
                case VnaSetUpMode.b2_b1_p2:
                    parameters = new VnaParameter[] { VnaParameter.Ratio(VnaReceiverType.b, 2, VnaReceiverType.b, 1, 2) };
                    this.disp4p = false;
                    this.visa.SrqWrite("SINC;USER4;DRIVPORT2;LOCKA2;NUMEB2;DENOB1;" +
                                       "CONVS;PARL \"b2/b1_p2\";REDD;ENTO");
                    break;
                case VnaSetUpMode.a1_p1:
                    parameters = new VnaParameter[] { VnaParameter.Receiver(VnaReceiverType.a, 1, 1) };
                    this.disp4p = false;
                    // Wiltron 360B if a1 is selected, no DENOMINATOR is allowed
                    this.visa.SrqWrite("SINC;USER1;DRIVPORT1;LOCKA1;DENONOR;NUMEA1;" +
                                       "CONVS;PARL \"a1_p1\";REDD;ENTO");
                    break;
                case VnaSetUpMode.b1_p1:
                    parameters = new VnaParameter[] { VnaParameter.Receiver(VnaReceiverType.b, 1, 1) };
                    this.disp4p = false;
                    this.visa.SrqWrite("SINC;USER2;DRIVPORT1;LOCKA1;DENONOR;NUMEB1;" +
                                       "CONVS;PARL \"b1_p1\";REDD;ENTO");
                    break;
                case VnaSetUpMode.b2_p1:
                    parameters = new VnaParameter[] { VnaParameter.Receiver(VnaReceiverType.b, 2, 1) };
                    this.disp4p = false;
                    this.visa.SrqWrite("SINC;USER3;DRIVPORT1;LOCKA1;DENONOR;NUMEB2;" +
                                       "CONVS;PARL \"b2_p1\";REDD;ENTO");
                    break;
                case VnaSetUpMode.a2_p1:
                    parameters = new VnaParameter[] { VnaParameter.Receiver(VnaReceiverType.a, 2, 1) };
                    this.disp4p = false;
                    this.visa.SrqWrite("SINC;USER4;DRIVPORT1;LOCKA1;DENONOR;NUMEA2;" +
                                       "CONVS;PARL \"a2_p1\";REDD;ENTO");
                    break;
                case VnaSetUpMode.a1_p2:
                    parameters = new VnaParameter[] { VnaParameter.Receiver(VnaReceiverType.a, 1, 2) };
                    this.disp4p = false;
                    // Wiltron 360B if a1 is selected, no DENOMINATOR is allowed
                    this.visa.SrqWrite("SINC;USER1;DRIVPORT2;LOCKA2;DENONOR;NUMEA1;" +
                                       "CONVS;PARL \"a1_p2\";REDD;ENTO");
                    break;
                case VnaSetUpMode.b1_p2:
                    parameters = new VnaParameter[] { VnaParameter.Receiver(VnaReceiverType.b, 1, 2) };
                    this.disp4p = false;
                    this.visa.SrqWrite("SINC;USER2;DRIVPORT2;LOCKA2;DENONOR;NUMEB1;" +
                                       "CONVS;PARL \"b1_p2\";REDD");
                    break;
                case VnaSetUpMode.b2_p2:
                    parameters = new VnaParameter[] { VnaParameter.Receiver(VnaReceiverType.b, 2, 2) };
                    this.disp4p = false;
                    this.visa.SrqWrite("SINC;USER3;DRIVPORT2;LOCKA2;DENONOR;NUMEB2;" +
                                       "CONVS;PARL \"b2_p2\";REDD;ENTO");
                    break;
                case VnaSetUpMode.a2_p2:
                    parameters = new VnaParameter[] { VnaParameter.Receiver(VnaReceiverType.a, 2, 2) };
                    this.disp4p = false;
                    this.visa.SrqWrite("SINC;USER4;DRIVPORT2;LOCKA2;DENONOR;NUMEA2;" +
                                       "CONVS;PARL \"a2_p2\";REDD;ENTO");
                    break;
                case VnaSetUpMode.xx_p1:
                    parameters = new VnaParameter[] { VnaParameter.Receiver(VnaReceiverType.a, 1, 1),
                                                      VnaParameter.Receiver(VnaReceiverType.b, 1, 1),
                                                      VnaParameter.Receiver(VnaReceiverType.b, 2, 1),
                                                      VnaParameter.Receiver(VnaReceiverType.a, 2, 1) };
                    this.disp4p = true;
                    // Wiltron 360B if a1 is selected, no DENOMINATOR is allowed
                    this.visa.SrqWrite("FOUPSPLI;USER1;USER2;USER3;USER4;ENTO");
                    this.visa.SrqWrite("USER1;DRIVPORT1;LOCKA1;DENONOR;NUMEA1;" +
                                       "CONVS;PARL \"a1_p1\";REDD");
                    this.visa.SrqWrite("USER2;DRIVPORT1;LOCKA1;DENONOR;NUMEB1;" +
                                       "CONVS;PARL \"b1_p1\";REDD");
                    this.visa.SrqWrite("USER3;DRIVPORT1;LOCKA1;DENONOR;NUMEB2;" +
                                       "CONVS;PARL \"b2_p1\";REDD");
                    this.visa.SrqWrite("USER4;DRIVPORT1;LOCKA1;DENONOR;NUMEA2;" +
                                       "CONVS;PARL \"a2_p1\";REDD");
                    break;
                case VnaSetUpMode.xx_p2:
                    parameters = new VnaParameter[] { VnaParameter.Receiver(VnaReceiverType.a, 1, 2),
                                                      VnaParameter.Receiver(VnaReceiverType.b, 1, 2),
                                                      VnaParameter.Receiver(VnaReceiverType.b, 2, 2),
                                                      VnaParameter.Receiver(VnaReceiverType.a, 2, 2) };
                    this.disp4p = true;
                    // Wiltron 360B if a1 is selected, no DENOMINATOR is allowed
                    this.visa.SrqWrite("FOUPSPLI;USER1;USER2;USER3;USER4;ENTO");
                    this.visa.SrqWrite("USER1;DRIVPORT2;LOCKA2;DENONOR;NUMEA1;" +
                                       "CONVS;PARL \"a1_p2\";REDD");
                    this.visa.SrqWrite("USER2;DRIVPORT2;LOCKA2;DENONOR;NUMEB1;" +
                                       "CONVS;PARL \"b1_p2\";REDD");
                    this.visa.SrqWrite("USER3;DRIVPORT2;LOCKA2;DENONOR;NUMEB2;" +
                                       "CONVS;PARL \"b2_p2\";REDD");
                    this.visa.SrqWrite("USER4;DRIVPORT2;LOCKA2;DENONOR;NUMEA2;" +
                                       "CONVS;PARL \"a2_p2\";REDD");
                    break;
                case VnaSetUpMode.xx:
                    parameters = new VnaParameter[] { VnaParameter.SParameter(1, 1),
                                                      VnaParameter.SParameter(2, 1),
                                                      VnaParameter.SParameter(1, 2),
                                                      VnaParameter.SParameter(2, 2),
                                                      VnaParameter.Receiver(VnaReceiverType.a, 1, 1),
                                                      VnaParameter.Receiver(VnaReceiverType.b, 1, 1),
                                                      VnaParameter.Receiver(VnaReceiverType.b, 2, 1),
                                                      VnaParameter.Receiver(VnaReceiverType.a, 2, 1),
                                                      VnaParameter.Receiver(VnaReceiverType.a, 1, 2),
                                                      VnaParameter.Receiver(VnaReceiverType.b, 1, 2),
                                                      VnaParameter.Receiver(VnaReceiverType.b, 2, 2),
                                                      VnaParameter.Receiver(VnaReceiverType.a, 2, 2) };
                    throw new System.Exception("VNA SetUp Mode xx not supported.");
            }
            setupmode = mode;
        }

        /// <summary>
        /// Set up Vna.
        /// </summary>
        /// <param name="mode">Vna mode</param>
        public void SetUp(string mode)
        {
            SetUp((VnaSetUpMode)System.Enum.Parse(typeof(VnaSetUpMode), mode));
        }

        /// <summary>
        /// Vna sweep mode: trigger hold.
        /// </summary>
        public void TriggerHold()
        {
            this.visa.SrqWrite("HLD");
        }

        int srq_timeout;

        /// <summary>
        /// Tiggers VNA.
        /// </summary>
        public void TriggerSingleStart()
        {
            srq_timeout = this.visa.SrqTimeout;
            this.visa.SrqMask = "CLES;SRQM 16";
            this.visa.SrqTimeout = 7200000;
            this.visa.SrqBegin();
            this.visa.Write("SING");
        }

        /// <summary>
        /// Waits for sweep complete.
        /// </summary>
        /// <param name="worker">Background Worker</param>
        /// <param name="e">Do Work Event Arguments</param>
        public void TriggerSingleWait(BackgroundWorker worker, DoWorkEventArgs e)
        {
            this.visa.SrqEnd(worker, e);
            this.visa.Clear();
            //this.visa.Write("CLES;SRQM 0;ENTO");
            this.visa.SrqMask = "CSB SRQ 128";
            this.visa.SrqTimeout = srq_timeout;
        }

        /// <summary>
        /// Waits for sweep complete.
        /// </summary>
        public void TriggerSingleWait()
        {
            VnaHelper.VnaTriggerSingleWait(this);
        }

        /// <summary>
        /// Triggers VNA and waits for sweep complete.
        /// </summary>
        /// <param name="worker">Background Worker</param>
        /// <param name="e">Do Work Event Arguments</param>
        public void TriggerSingle(BackgroundWorker worker, DoWorkEventArgs e)
        {
            TriggerSingleStart();
            TriggerSingleWait(worker, e);
        }

        /// <summary>
        /// Triggers VNA and waits for sweep complete.
        /// </summary>
        public void TriggerSingle()
        {
            TriggerSingleStart();
            TriggerSingleWait();
        }

        /// <summary>
        /// Vna sweep mode: trigger continuos.
        /// </summary>
        public void TriggerCont()
        {
            this.visa.SrqWrite("CONT");
        }

        /// <summary>
        /// Get data from Vna. All available parameters.
        /// </summary>
        /// <param name="format">Vna format</param>
        /// <returns>Data</returns>
        public VnaData GetData(VnaFormat format)
        {
            return GetData(Parameters, format);
        }

        /// <summary>
        /// Get data from vna.
        /// </summary>
        /// <param name="parameters">Vna parameters</param>
        /// <param name="format">Vna format</param>
        /// <returns>Data</returns>
        private VnaData GetData(VnaParameter[] parameters, VnaFormat format)
        {
            double[] temp;
            double[] flist = GetFrequencyList();
            var ports = VnaParameter.CommonPorts(parameters);
            double z0 = System.Math.Round(Z0, 3);
            int n1 = flist.Length;
            int n2 = parameters.Length;
            VnaData data = new VnaData();
            data.Frequency = flist;
            data.Ports = ports;
            data.PortZr = new Complex<Number>[ports.Length];
            for (int i1 = 0; i1 < ports.Length; i1++)
            {
                data.PortZr[i1] = z0;
            }
            data.ParameterData = new VnaParameterData<Number>[n2];
            for (int i2 = 0; i2 < n2; i2++)
            {
                byte[] bindata;
                bool conv1s = false;
                string wb = "";
                byte x;
                switch (parameters[i2].Name)
                {
                    case "S1,1":
                        if (format == VnaFormat.RawData)
                        {
                            if (disp4p) x = 1; else x = 1;
                            wb = "S11;FORM3;OUTPRAW" + x.ToString();
                        }
                        else
                        {
                            wb = "S11;FORM3;OUTPDATA";
                        }
                        break;
                    case "S2,1":
                        if (format == VnaFormat.RawData)
                        {
                            if (disp4p) x = 2; else x = 1;
                            wb = "S21;FORM3;OUTPRAW" + x.ToString();
                        }
                        else
                        {
                            wb = "S21;FORM3;OUTPDATA";
                        }
                        break;
                    case "S1,2":
                        if (format == VnaFormat.RawData)
                        {
                            if (disp4p) x = 3; else x = 1;
                            wb = "S12;FORM3;OUTPRAW" + x.ToString();
                        }
                        else
                        {
                            wb = "S12;FORM3;OUTPDATA";
                        }
                        break;
                    case "S2,2":
                        if (format == VnaFormat.RawData)
                        {
                            if (disp4p) x = 4; else x = 1;
                            wb = "S22;FORM3;OUTPRAW" + x.ToString();
                        }
                        else
                        {
                            wb = "S22;FORM3;OUTPDATA";
                        }
                        break;
                    case "b1/b2_p1":
                        if (disp4p) x = 1; else x = 1;
                        conv1s = true;
                        wb = "USER1;FORM3;OUTPRAW" + x.ToString();
                        break;
                    case "b2/b1_p2":
                        if (disp4p) x = 4; else x = 1;
                        wb = "USER4;FORM3;OUTPRAW" + x.ToString();
                        break;
                    case "a1_p1":
                        if (disp4p) x = 1; else x = 1;
                        wb = "USER1;FORM3;OUTPRAW" + x.ToString();
                        break;
                    case "b1_p1":
                        if (disp4p) x = 2; else x = 1;
                        wb = "USER2;FORM3;OUTPRAW" + x.ToString();
                        break;
                    case "b2_p1":
                        if (disp4p) x = 3; else x = 1;
                        wb = "USER3;FORM3;OUTPRAW" + x.ToString();
                        break;
                    case "a2_p1":
                        if (disp4p) x = 4; else x = 1;
                        wb = "USER4;FORM3;OUTPRAW" + x.ToString();
                        break;
                    case "a1_p2":
                        if (disp4p) x = 1; else x = 1;
                        wb = "USER1;FORM3;OUTPRAW" + x.ToString();
                        break;
                    case "b1_p2":
                        if (disp4p) x = 2; else x = 1;
                        wb = "USER2;FORM3;OUTPRAW" + x.ToString();
                        break;
                    case "b2_p2":
                        if (disp4p) x = 3; else x = 1;
                        wb = "USER3;FORM3;OUTPRAW" + x.ToString();
                        break;
                    case "a2_p2":
                        if (disp4p) x = 4; else x = 1;
                        wb = "USER4;FORM3;OUTPRAW" + x.ToString();
                        break;
                    case "a1/b1_p2":
                        if (disp4p) x = 1; else x = 1;
                        conv1s = true;
                        wb = "USER1;FORM3;OUTPRAW" + x.ToString();
                        break;
                    case "a2/b2_p1":
                        if (disp4p) x = 4; else x = 1;
                        conv1s = true;
                        wb = "USER4;FORM3;OUTPRAW" + x.ToString();
                        break;
                }
                this.visa.SrqWrite(wb);
                bindata = this.visa.ReadByteArray();
                // form3
                temp = VnaHelper.Form3ByteArray2DoubleArray(bindata);
                // Vna data
                data.ParameterData[i2] = new VnaParameterData<Number>();
                data.ParameterData[i2].Parameter = parameters[i2];
                data.ParameterData[i2].Data = new Complex<Number>[n1];
                for (int i1 = 0; i1 < n1; i1++)
                {
                    if (conv1s)
                        data.ParameterData[i2].Data[i1] = 1 / new Complex<Number>(temp[2 * i1], temp[2 * i1 + 1]);
                    else 
                        data.ParameterData[i2].Data[i1] = new Complex<Number>(temp[2 * i1], temp[2 * i1 + 1]);
                }
            }
            return data;
        }

        /// <summary>
        /// Triggers Vna, waits for sweep complete, gets and saves raw data.
        /// </summary>
        /// <param name="pathRes">Data File Path (*.sdatb, *.vdatb)</param>
        public void Measure(string pathRes)
        {
            VnaHelper.VnaMeasure(this, pathRes);
        }

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
        /// Instrument Identification
        /// </summary>
        public string Identification
        {
            get
            {
                return this.visa.Query("OUTPIDEN").Split('\n')[0];
            }
        }

        /// <summary>
        /// Settings
        /// </summary>
        public VnaSettings Settings
        {
            get { return VnaHelper.GetVnaSettings(this); }
        }

        /// <summary>
        /// Sweep Mode
        /// </summary>
        public VnaSweepMode SweepMode
        {
            get
            {
                VnaSweepMode value = VnaSweepMode.LinearFrequency;
                int t1 = this.visa.Timeout;
                int t2 = this.visa.SrqTimeout;
                this.visa.Timeout = 20000;
                this.visa.SrqTimeout = 20000;
                string mode = this.visa.SrqQuery("SWEM?");
                this.visa.Timeout = t1;
                this.visa.SrqTimeout = t2;
                mode = mode.Split('"')[1];
                switch (mode)
                {
                    case "RAMP":
                        value = VnaSweepMode.LinearFrequency;
                        break;
                    case "STEP":
                        value = VnaSweepMode.LinearFrequency;
                        break;
                    case "FREQUENCY  LIST":
                        value = VnaSweepMode.SegmentSweep;
                        break;
                    case "SINGLE POINT":
                        value = VnaSweepMode.CWTime;
                        break;
                    case "FAST CW":
                        value = VnaSweepMode.CWTime;
                        break;
                }
                return value;
            }
            set
            {
                string mode = "";
                switch (value)
                {
                    case VnaSweepMode.LinearFrequency:
                        mode = "STEP";
                        break;
                    case VnaSweepMode.LogFrequency:
                        throw new System.Exception("VNA Sweep Mode LogFrequency not supported.");
                    case VnaSweepMode.SegmentSweep:
                        // workaround for empty segment table
                        int nsegm = VnaHelper.String2Int(this.visa.SrqQuery("NUMS?"));
                        if (nsegm == 0)
                        {
                            double start = this.FrequencyStart;
                            double stop = this.FrequencyStop;
                            double step = (stop - start) / (this.SweepPoints - 1);
                            this.SegmentTable = new double[,] 
                            { { start, stop, step, 0 } };
                        }
                        mode = "LISFREQ;ASEG";
                        break;
                    case VnaSweepMode.CWTime:
                        mode = "SINP";
                        break;
                }
                this.visa.SrqWrite(mode);
            }
        }

        /// <summary>
        /// Sweep Time (s)
        /// </summary>
        public double SweepTime
        {
            get
            {
                return VnaHelper.String2Double(this.visa.SrqQuery("SWET;OUTPACTI"));
            }
            set
            {
                this.visa.SrqWrite("SWET " + value.ToString("E12", CultureInfo.InvariantCulture));
            }
        }

        /// <summary>
        /// Dwell Time (s)
        /// </summary>
        public double DwellTime
        {
            get
            {
                return VnaHelper.String2Double(this.visa.SrqQuery("DWET;OUTPACTI"));
            }
            set
            {
                this.visa.SrqWrite("DWET " + value.ToString("E12", CultureInfo.InvariantCulture));
            }
        }

        /// <summary>
        /// Sweep Points
        /// </summary>
        public int SweepPoints
        {
            get
            {
                return VnaHelper.String2Int(this.visa.SrqQuery("POIN;OUTPACTI"));
            }
            set
            {
                this.visa.SrqWrite("POIN " + value.ToString());
            }
        }

        /// <summary>
        /// IF Bandwidth (Hz)
        /// </summary>
        public double IFBandwidth
        {
            get
            {
                return 0;
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
                int state = VnaHelper.String2Int(this.visa.SrqQuery("AVER?"));
                int value = 1;
                if (state == 1)
                {
                    value = VnaHelper.String2Int(this.visa.SrqQuery("AVERON;OUTPACTI"));
                }
                return value;
            }
            set
            {
                if (value <= 1)
                {
                    this.visa.SrqWrite("AVEROFF");
                }
                else
                {
                    this.visa.SrqWrite("AVERON " + value.ToString());
                }
            }
        }

        /// <summary>
        /// Start Frequency (Hz)
        /// </summary>
        public double FrequencyStart
        {
            get
            {
                return VnaHelper.String2Double(this.visa.SrqQuery("STAR;OUTPACTI"));
            }
            set
            {
                this.visa.SrqWrite("STAR " + value.ToString("E12", CultureInfo.InvariantCulture));
            }
        }

        /// <summary>
        /// Stop Frequency (Hz)
        /// </summary>
        public double FrequencyStop
        {
            get
            {
                return VnaHelper.String2Double(this.visa.SrqQuery("STOP;OUTPACTI"));
            }
            set
            {
                this.visa.SrqWrite("STOP " + value.ToString("E12", CultureInfo.InvariantCulture));
            }
        }

        /// <summary>
        /// Center Frequency (Hz)
        /// </summary>
        public double FrequencyCenter
        {
            get
            {
                return VnaHelper.String2Double(this.visa.SrqQuery("CENT;OUTPACTI"));
            }
            set
            {
                this.visa.SrqWrite("CENT " + value.ToString("E12", CultureInfo.InvariantCulture));
            }
        }

        /// <summary>
        /// Span Frequency (Hz)
        /// </summary>
        public double FrequencySpan
        {
            get
            {
                return VnaHelper.String2Double(this.visa.SrqQuery("SPAN;OUTPACTI"));
            }
            set
            {
                this.visa.SrqWrite("SPAN " + value.ToString("E12", CultureInfo.InvariantCulture));
            }
        }

        /// <summary>
        /// CW Frequency (Hz)
        /// </summary>
        public double FrequencyCW
        {
            get
            {
                return this.FrequencyCenter;
            }
            set
            {
                this.FrequencyCenter = value;
            }
        }

        /// <summary>
        /// SegmentTable
        /// 1. Column : Start Frequency (Hz)
        /// 2. Column : Stop Frequency (Hz)
        /// 3. Column : Frequency Step (Hz)
        /// 4. Column : IF Bandwidth (Hz)
        /// </summary>
        public double[,] SegmentTable
        {
            get
            {
                double start, stop, step, ifbw;
                int n = VnaHelper.String2Int(this.visa.SrqQuery("NUMS?"));
                double[,] value = new double[n, 4];
                for (int i = 0; i < n; i++)
                {
                    this.visa.SrqWrite("SEDI" + (i + 1).ToString());
                    start = VnaHelper.String2Double(this.visa.SrqQuery("STAR;OUTPACTI"));
                    stop = VnaHelper.String2Double(this.visa.SrqQuery("STOP;OUTPACTI"));
                    step = VnaHelper.String2Double(this.visa.SrqQuery("STPSIZE;OUTPACTI"));
                    ifbw = this.IFBandwidth;
                    value[i, 0] = start;
                    value[i, 1] = stop;
                    value[i, 2] = step;
                    value[i, 3] = ifbw;
                }
                return value;
            }
            set
            {
                VnaSweepMode sweep_mode = this.SweepMode;
                int t1 = this.visa.Timeout;
                int t2 = this.visa.SrqTimeout;
                this.visa.Timeout = 20000;
                this.visa.SrqTimeout = 20000;
                this.visa.SrqWrite("EDITLIST");
                this.visa.SrqWrite("CLEL");
                double start, stop, step, ifbw;
                int n = value.GetLength(0);
                for (int i = 0; i < n; i++)
                {
                    start = value[i, 0];
                    stop = value[i, 1];
                    step = value[i, 2];
                    ifbw = value[i, 3];
                    this.visa.SrqWrite("SADD");
                    this.visa.SrqWrite("STAR " + start.ToString("E12", CultureInfo.InvariantCulture));
                    this.visa.SrqWrite("STOP " + stop.ToString("E12", CultureInfo.InvariantCulture));
                    this.visa.SrqWrite("STPSIZE " + step.ToString("E12", CultureInfo.InvariantCulture));
                    this.visa.SrqWrite("SDON");
                }
                this.visa.SrqWrite("DUPD");
                this.visa.SrqWrite("EDITDONE");
                this.visa.Timeout = t1;
                this.visa.SrqTimeout = t2;
                this.SweepMode = sweep_mode;
            }
        }

        /// <summary>
        /// Source 1 Power (dBm)
        /// </summary>
        public double Source1Power
        {
            get
            {
                return VnaHelper.String2Double(this.visa.SrqQuery("POWE;OUTPACTI"));
            }
            set
            {
                this.visa.SrqWrite("POWE " + value.ToString("E12", CultureInfo.InvariantCulture));
            }
        }

        /// <summary>
        /// Source 1 Power Slope (dB/GHz)
        /// </summary>
        public double Source1PowerSlope
        {
            get
            {
                int state = VnaHelper.String2Int(this.visa.SrqQuery("SLOP?"));
                int value = 0;
                if (state == 1)
                {
                    value = VnaHelper.String2Int(this.visa.SrqQuery("SLOPON;OUTPACTI"));
                }
                return value;
            }
            set
            {
                if (value <= 0)
                {
                    this.visa.SrqWrite("SLOPOFF");
                }
                else
                {
                    this.visa.SrqWrite("SLOPON " + value.ToString());
                }
            }
        }

        /// <summary>
        /// Port 1 Attenuator (dB)
        /// </summary>
        public double Port1Attenuator
        {
            get
            {
                return VnaHelper.String2Double(this.visa.SrqQuery("ATTP1;OUTPACTI"));
            }
            set
            {
                this.visa.SrqWrite("ATTP1 " + value.ToString("E12", CultureInfo.InvariantCulture));
            }
        }

        /// <summary>
        /// Port 1 Extension (s)
        /// </summary>
        public double Port1Extension
        {
            get
            {
                return VnaHelper.String2Double(this.visa.SrqQuery("PORT1;OUTPACTI"));
            }
            set
            {
                this.visa.SrqWrite("PORT1 " + value.ToString("E12", CultureInfo.InvariantCulture));
            }
        }

        /// <summary>
        /// Source 2 Power (dBm)
        /// </summary>
        public double Source2Power
        {
            get
            {
                return VnaHelper.String2Double(this.visa.SrqQuery("POW2;OUTPACTI"));
            }
            set
            {
                this.visa.SrqWrite("POW2 " + value.ToString("E12", CultureInfo.InvariantCulture));
            }
        }

        /// <summary>
        /// Source 2 Power Slope (dB/GHz)
        /// </summary>
        public double Source2PowerSlope
        {
            get
            {
                int state = VnaHelper.String2Int(this.visa.SrqQuery("SLOP2?"));
                int value = 0;
                if (state == 1)
                {
                    value = VnaHelper.String2Int(this.visa.SrqQuery("SLOP2ON;OUTPACTI"));
                }
                return value;
            }
            set
            {
                if (value <= 0)
                {
                    this.visa.SrqWrite("SLOP2OFF");
                }
                else
                {
                    this.visa.SrqWrite("SLOP2ON " + value.ToString());
                }
            }
        }

        /// <summary>
        /// Port 2 Attenuator (dB)
        /// </summary>
        public double Port2Attenuator
        {
            get
            {
                return VnaHelper.String2Double(this.visa.SrqQuery("ATTP2;OUTPACTI"));
            }
            set
            {
                this.visa.SrqWrite("ATTP2 " + value.ToString("E12", CultureInfo.InvariantCulture));
            }
        }

        /// <summary>
        /// Port 2 Extension (s)
        /// </summary>
        public double Port2Extension
        {
            get
            {
                return VnaHelper.String2Double(this.visa.SrqQuery("PORT2;OUTPACTI"));
            }
            set
            {
                this.visa.SrqWrite("PORT2 " + value.ToString("E12", CultureInfo.InvariantCulture));
            }
        }

        private double _z0 = 0;

        /// <summary>
        /// Z0 (Ohm)
        /// </summary>
        public double Z0
        {
            get
            {
                if (_z0 == 0)
                    _z0 = VnaHelper.String2Double(this.visa.SrqQuery("SETZ;OUTPACTI"));
                return _z0;
            }
            set
            {
                _z0 = value;
                this.visa.SrqWrite("SETZ " + value.ToString("E12", CultureInfo.InvariantCulture));
            }
        }

        /// <summary>
        /// Turns RF power from the source ON or OFF.
        /// </summary>
        public bool OutputState
        {
            get
            {
                return true;
            }
            set
            {
                throw new System.NotImplementedException();
            }
        }

        private VnaParameter[,] parameterMatrix = null;

        /// <summary>
        /// Vna parameter matrix
        /// </summary>
        public VnaParameter[,] ParameterMatrix
        {
            get
            {
                return parameterMatrix;
            }
            set
            {
                SetUp(VnaHelper.GetVnaSetUpMode(value));
                parameterMatrix = value;
            }
        }

        /// <summary>
        /// Vna Set up mode
        /// </summary>
        public VnaSetUpMode SetUpMode
        {
            get
            {
                return this.setupmode;
            }
        }

        /// <summary>
        /// List with available parameters
        /// </summary>
        public VnaParameter[] Parameters
        {
            get
            {
                return this.parameters;
            }
        }

        #endregion

        #region Private Methods

        private double[] GetFrequencyList()
        {
            int points;
            double start;
            double stop;
            double[] frequency_list = new double[0];
            VnaSweepMode sweep_mode = this.SweepMode;
            switch (sweep_mode)
            {
                case VnaSweepMode.LinearFrequency:
                    points = this.SweepPoints;
                    start = this.FrequencyStart;
                    stop = this.FrequencyStop;
                    frequency_list = NumLib.Linspace(start, stop, points);
                    break;
                case VnaSweepMode.LogFrequency:
                    points = this.SweepPoints;
                    start = this.FrequencyStart;
                    stop = this.FrequencyStop;
                    frequency_list = NumLib.Logspace(System.Math.Log10(start), System.Math.Log10(stop), points);
                    break;
                case VnaSweepMode.SegmentSweep:
                    double[,] table = this.SegmentTable;
                    int n = table.GetLength(0);
                    for (int i = 0; i < n; i++)
                    {
                        start = table[i, 0];
                        stop = table[i, 1];
                        points = (int)System.Math.Floor((stop - start) / table[i, 2] + 1);
                        double[] temp_list = new double[frequency_list.Length + points];
                        frequency_list.CopyTo(temp_list, 0);
                        NumLib.Linspace(start, stop, points).CopyTo(temp_list, frequency_list.Length);
                        frequency_list = temp_list;
                    }
                    break;
                case VnaSweepMode.CWTime:
                    points = this.SweepPoints;
                    start = this.FrequencyCW;
                    stop = start;
                    frequency_list = NumLib.Linspace(start, stop, points);
                    break;
            }
            return frequency_list;
        }

        #endregion

    }
}
