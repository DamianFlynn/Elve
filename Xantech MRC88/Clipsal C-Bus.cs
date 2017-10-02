using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using CodecoreTechnologies.Elve.DriverFramework;
using CodecoreTechnologies.Elve.DriverFramework.Scripting;
using CodecoreTechnologies.Elve.DriverFramework.DeviceSettingEditors;
using CodecoreTechnologies.Elve.DriverFramework.Communication;
using System.Timers;
using CodecoreTechnologies.Elve.DriverFramework.DriverInterfaces;
using CodecoreTechnologies.Elve.DriverFramework.Extensions;
using System.Xml.Linq;

namespace BuiltInDrivers
{

    /// <summary>
    /// 
    /// Manuals:
    /// http://training.clipsal.com/downloads/OpenCBus/Serial%20Interface%20User%20Guide.pdf
    ///
    /// Overview:
    /// This driver is used to control and monitor the Clipsal C-Bus network using the serial interface. The driver is tested using a C-BUS 5500 PC Interface
    /// 
    /// Driver Lifecycle:
    ///   1. The driver is instantiated by the associated driver service.
    ///   2. The driver's optional ConfigurationFileNames property is read by the driver service so it can later provide the driver with any extra configuration file data.
    ///   3. Driver parameters that have been configured by the user are set. (Those properties with the DriverParameter attribute.)
    ///   4. StartDriver() is called, which includes any extra configuration file data as a parameter.
    ///   5. Driver events collections are populated with configured rules. (Those fields or properties with the DriverEvent attribute.)
    ///   6. The driver's optional InitializeRules() method is called.
    ///   7. StopDriver() is called when the driver is to be stopped.
    ///   
    /// </summary>


    [DriverAttribute(
        "Clipsal C-Bus Driver",                                                                        // Display Name 
        "This driver supports control and monitoring of the Clipsal C-Bus Network. " +                  // Description
        "This driver is coded to interface with the Lighting Application, and to interface " +
        "to the C-Bus network using the 5500PC module via RS232. This protocol is locked to " +
        "9600 Baud, and the driver we reset the interface at startup",
        "Damian Flynn",                                                                                 // Author
        "Lighting & Electrical",                                                                        // Category
        "",                                                                                             // Subcategory
        "CBus" ,                                                                                        // Default instance name
        DriverCommunicationPort.Serial,                                                                 // Default Communications Metod
        DriverMultipleInstances.MultiplePerDriverService,                                               // Allow multiple instances
        0,                                                                                              // Major Version
        40,                                                                                             // Minor Version
        DriverReleaseStages.Development,                                                                // Release Stage
        "Clipsal",                                                                                      // Manufacturer Name
        "http://www.clipsal.com",                                                                       // Manufacturer ULR
        "MxFSE98SDsiNQPeacuSfIwZhbR4Q2Bhdv2DqX/IBRJvf7CeBvrVneg=="                                      // Registration - for offical use
        )]


    class ClipsalCBusDriver : Driver, ILightingAndElectricalDriver
    {

        #region Variables

        // Define Communications Protocol Variables
        private SerialCommunication _serial;
        private string _serialPortName;                 // "COM1"
        private const string TX_TERMINATOR = "\r";
        private const string RX_TERMINATOR = "\r"; // this is the incoming message delimiter

        // Define a timer object as we will be executing peroiodic status checks
        private System.Timers.Timer _timer;


        // Driver Variables

        // Hardware Physical Limits

        private const int MaxZoneCount = 256;        // Number of Lighting Zones

        // Zone Details
        private int _configuredZoneCount;

        private string[] _lightNames = new string[MaxZoneCount];
        private bool[] _lightPowerStates = new bool[MaxZoneCount];
        private int[] _lightPowerLevels = new int[MaxZoneCount];

        // Ramp Rates
        public Dictionary<string, byte> RampRates;

        #endregion

        /// <summary>
        /// Initialization of the driver to default values
        /// </summary>
        public ClipsalCBusDriver()
        {
            // Initialize zone related vars            
            for (int i = 0; i < MaxZoneCount; i++)
            {
                _lightNames[i] = "Light Group " + (i);  // C-Bus indexes all Light Groups from 0 to 255 (00-FF), However recommend that Group 0 is not actually used.
                _lightPowerStates[i] = false;
                _lightPowerLevels[i] = -1;
            }

            // Define Settings applicable for the driver
            _configuredZoneCount = MaxZoneCount;

            // Setup a Key Value Pair to Define the Ramp Rates
            RampRates = new Dictionary<string, byte>();

            RampRates.Add("0 Sec", 0x02);
            RampRates.Add("4 Sec", 0x0A);
            RampRates.Add("8 Sec", 0x12);
            RampRates.Add("12 Sec", 0x1A);
            RampRates.Add("20 Sec", 0x22);
            RampRates.Add("30 Sec", 0x2A);
            RampRates.Add("40 Sec", 0x32);
            RampRates.Add("60 Sec", 0x3A);
            RampRates.Add("90 Sec", 0x42);
            RampRates.Add("2 Min", 0x4A);
            RampRates.Add("3 Min", 0x52);
            RampRates.Add("5 Min", 0x5A);
            RampRates.Add("7 Min", 0x62);
            RampRates.Add("10 Min", 0x6A);
            RampRates.Add("15 Min", 0x72);
            RampRates.Add("17 Min", 0x7A);
            RampRates.Add("Stop Ramp", 0x09);
        }



        #region DriverSettingAttribute
        // Provide setting for the driver to be presented in the Elve managment studio


        /// <summary>
        /// Communcation Port we will Utilize for Communications
        /// </summary>
        [DriverSettingAttribute("Serial Port Name", "The name of the serial port that C-Bus is connected to. Ex. COM1", typeof(SerialPortDeviceSettingEditor), null, true)]
        public string SerialPortNameSetting
        {
            set { _serialPortName = value; }
        }


        /// <summary>
        /// Assign a friendly name for each zone on the Matrix, Starting at Group 1 trought to 255.
        /// Note that group 0 is masked here, as this group is recommended NOT to be assigned in the C-Bus toolkit.
        /// </summary>
        [DriverSettingArrayNamesAttribute("Light Names", "Enter the name of each Light Zone output on the Matrix.", typeof(ArrayItemsDriverSettingEditor), "ZoneNames", 0, MaxZoneCount, "", false)]
        public string CustomLightNames
        {
            set
            {
                if (string.IsNullOrEmpty(value) == false)
                {
                    XElement element = XElement.Parse(value);
                    // This uses LINQ... junkNames is not really used but ToList() needs to be called to force the query to run.
                    List<string> junkNames = (
                        from node in element.Elements("Item")
                        select _lightNames[(int)node.Attribute("Index")] = node.Attribute("Name").Value).ToList();
                }
            }
        }


        #endregion

        #region Core Driver Function

        /// <summary>
        /// Starts the driver.  This typcially sets any class variables and hooks
        /// any event handlers such as SerialPort ReceivedBytes, etc.
        /// </summary>
        /// <param name="configFileData">Contains the contents of any configuration files specified in the ConfigurationFileNames property.</param>
        public override bool StartDriver(Dictionary<string, byte[]> configFileData)
        {

            Logger.DebugFormat("{0}: Opening serial port.", this.DeviceName);


            // Communication settings should be 9600 baud, 8 data bits, no parity, and one stop bit.
            _serial = new SerialCommunication(_serialPortName, 9600, System.IO.Ports.Parity.None, 8, System.IO.Ports.StopBits.One, System.IO.Ports.Handshake.None,true,true);
            _serial.CurrentEncoding = new EightBitEncoding(); // we can't use ASCIIEncoding since it converts 0x80 and greater to '?'.
            _serial.Delimiter = RX_TERMINATOR; // this is the incoming message delimiter
            _serial.ReceivedDelimitedString += new EventHandler<ReceivedDelimitedStringEventArgs>(_serial_ResponseReceived);

            // Define the Logger Object for the Serial Device
            _serial.Logger = Logger; 

            // Set up connection monitor.
            _serial.ConnectionMonitorTimeout = 60000;           // ensure we receive data at least once every minute so we know the connection is alive
            _serial.ConnectionMonitorTestRequest = "";          // request version every 60 seconds to ensure the unit is still connected.
            _serial.ConnectionEstablished += new EventHandler<EventArgs>(_serial_ConnectionEstablished);
            _serial.ConnectionLost += new EventHandler<EventArgs>(_serial_ConnectionLost);
            _serial.StartConnectionMonitor();                   // this will also attempt to open the serial connection


            // set timer to get the current device state
            _timer = new System.Timers.Timer();
            _timer.AutoReset = false;
            _timer.Elapsed += new ElapsedEventHandler(_timer_Elapsed);
            _timer.Interval = 180000;
            _timer.Start();

            // As we need to gather information from the device, we will indicate that we are not yet in Ready State.
            return false;

        }

        /// <summary>
        /// Stops the driver by unhooking any event handlers and releasing any used resources.
        /// </summary>
        public override void StopDriver()
        {
            // TODO: Add any necessary cleanup logic here.
            if (_timer != null)
            {
                _timer.Dispose();
                _timer = null;
            }
            if (_serial != null)
            {
                _serial.Dispose();
                _serial = null;
            }
        }

        /// <summary>
        /// Returns a list of configuration file names that the driver requires.  The files will be read from the same folder that the master server configuration file exists in.
        /// If your driver needs configuration file(s), override this property.  By default this returns an empty array.
        /// </summary>
        public override string[] ConfigurationFileNames
        {
            get
            {
                return new string[] { };
            }
        }


        void _serial_ConnectionEstablished(object sender, EventArgs e)
        {
            Logger.DebugFormat("{0}: Opened Serial Connection. (This does not guarantee that the hardware is physically connected to the serial port).", this.DeviceName);

            
            //
            // Clipsal C-Bus RS232 Gateway Initialazation process
            //

            // Before we begin with Any initialzation procedure, we will send the reset command a few times to get the interface prepaired
            
            sendCommand("~~~");


            // ::Protocol Overview
            //
            // The protcol used to communicate with the C-Bus system is ASCII character based
            // Instuctions for commands are constucted from BIT level flags
            // These are stored in Byte blocks, to allow simpler checksum calculations
            //
            // :: Initialzation Commands
            //
            // The Serial to C-Bus interface is an intelligent device, and requires some initiazaiton commands to get the device
            // active and running in the correct mode for us to sucesfully use the gateway to communicate with the bus.
            // 
            // Initialization Command String Format:
            //    @A3pp00vv<cr> 
            //       pp = Command to be Issued
            //       vv = Value of the Command Payload 
            //       
            
            // :: Initialzation Sequence
            //
            // We will issue a series of commands to the interface to set it up in the correct mode for our usage
            //

            // : Stage 1 - Application Listening Selection
            //
            // Clipsal C-Bus is a multi purpose protcol interface, with support for may applications ranging from Lighting, Heating, Audio,
            // Secruity, etc. Due to the Vast amount of data which could be presented on the gateway, we have the option to subscribe to 
            // any two applications per gateway, or to Applications.
            //
            // The Command (pp) to be Issued is the "Application Address" and we only need one, so will just use "Application Address 1 ($21)"
            //
            // Since this driver is coded primary to integrate with the lighting application (Indexed as Application 56 (0x38) in C-Bus), we
            // will instruct the interface that we are going to subscribe to notifications for the application only.
            // Thus the Command Value (vv) we will use is $38 - which is Hex for C-Bus Application Number 56 (Lighting)
            //
            // In this command we will replace the paramaters of the command sting based on the following 
            //
            //    pp = $21 - Application Address 1 :: vv = $38|$FF|$?? (C-Bus Application we are subscribing to E.G. $38 = Lighting, $FF = ALL Applicaions)
            //    pp = $22 - Application Address 2 :: vv = $38|$FF|$?? (Repeated for the Second Application to subscribe, if App1 is $FF, this must be also)
            //

            sendCommand("@A3210038");  // Application Address 1 :: Request to Reciece Lighting Application Updates

            // : Stage 2 - Set Various Serial Interface Communications options 
            //
            // As we mentioned the Gateway is an intelligent device, there are 3 sets of options we can enable and disable options on
            // this set is documented as "Interface Options 3" and by default at power on are set to Basic Mode (all off - $00)
            // 
            // The Command (pp) to be Issued is the "Interface Options 3 - $A3" 
            //
            // The Command Value  (vv) is created from a binary list of 8 options, and converted to a hex number for transmission
            // The following is a look at each of the 8 settings, and their purpose
            // 
            //
            //   0. PCN        - When set, a parameter change notification (PCN) message will be emitted by the serial interface any time
            //                   it detects an attempt to update on or more of its internal paramaters from the C-Bus
            //
            //   1. LOCAL_SAL  - When Set the Serial Interface will
            //                   a) Issue Point to Point commands to the local network as local messages (empty network PCI)
            //                   b) Issue Point to Point to Multipoint commands as they they originated from the local network
            //                   This allows the Serial Interface to initiate Point to Multipoint commands to behave exactly like commands
            //                   issued by simple (key input) units. If this is not set the commands will have a slightly different format
            //                   which could cause minor incompatibility with some devices that do not accept these messages
            //
            //   2. PUN        - When Set, a Power Up Notification message is emitted by the serial interface after it has completed its init
            //                   and before accepting any charactes from its serial input port
            //
            //   3. EXSTAT     - When Set
            //                   a) Switched Status Replies are presented in a format compatible with the Level Staus Reply, and
            //                   b) All Status Replies are presented in a Long Form with Addressing Information
            //
            //   4. Reserved   - Set to 0
            //
            //   5. Reserved   - Set to 0
            //
            //   6. Reserved   - Set to 0
            //
            //   7. Reserved   - Set to 0
            //
            // Note: Clipsal has published a recommendation that all new devices using the Serial Interface Enable the option LOCAL_SAL
            // Based on these recommendations EXSTAT will also be enabled.
            //
            // We now get to select the options which we would like to enable on the gateway, and place them in the following sequence so that we can 
            // then convert this number from its Binary representation 10010100 to $FF Hex format to transmit to the gateway
            //
            //                 76543210   Flags
            //   vv = $06 ::   00000110 = LOCAL_SAL|EXSTAT
            //
            

            sendCommand("@A3420006");  // Interface options 3 :: Param: $42 Value: $06


            // : Stage 3 - Set Various Serial Interface Communications options (More!)
            //
            // This next set is documented as "Interface Options 1" and by default at power on are set to Basic Mode (all off - $00)
            // 
            // The Command (pp) to be Issued is the "Interface Options 1 - $30" 
            //
            // The Command Value  (vv) is created from a binary list of 8 options, and converted to a hex number for transmission
            // The following is a look at each of the 8 settings, and their purpose
            //
            //   0. CONNECT    - When set, the serial Interface connects the C-Bus network to the RS-232 port for applications configured with App Addr ($21 and $22)
            //                   This connection allows us to receive all C-Bus Point to Multipoint SAL messages for the defined applications
            //                   When the CONNECT option is only is set, Status reports are not passed to us. They can be enabled by setting the MONITOR option
            //
            //   1. Reserved   - Set to 0
            //
            //   2. XONXOFF    - When Set, switches on the use of XON/XOFF handshaking for RS-232 Comms.
            //                   If set, 
            //                     RS232 will issue an XON character once each time its serial buffer is emptied.
            //                     The XOFF character will be sent once when its buffer (20 bytes) reaches the 14th byte
            //
            //   3. SRCHK      - When Set, forces the Serial Interace to expect a checksum on all serial communciations it revieves.
            //                   If the serial interface detects a checksum error, command will be ignored and a ! will be returned
            //
            //   4. SMART      - When set, the Serial Interface will NOT echo serial data it receives, and will include all path information
            //                   and source address in the monitored SAL messages and some CAL Replies
            //                   Note - Full addressing is not included for CAL replies from itself and Status reports (Use IDMON and EXSTAT to turn these on)
            //  
            //   5. MONITOR    - When set, the Serial Interface will relay all Status Reports for Applications matching paramates in applications 1 and 2 ($21 and $22)
            //                   The Form of the status report will depend on the setting in EXSTAT
            //
            //   6. IDMON      - When set, all messages returned from the RS232 in respose to a command are given in a format consistent with SMART mode
            //   
            //   7. Reserved   - Must be set to 0
            //
            // Note: Clipsal has published a recommendation on the combination of these options to make the processing of the status messages simpler
            // the suggested combination is what I am going to select from these options
            //
            // We now get to select the options which we would like to enable on the gateway, and place them in the following sequence so that we can 
            // then convert this number from its Binary representation 10010100 to $FF Hex format to transmit to the gateway
            //
            //                 76543210   Flags
            //   vv = $79 ::   01111001 = CONNECT|SRCHK|SMART|MONITOR|IDMON
            //
            
            sendCommand("@A3300079");  // Interface options 1 :: Param: $30 Value: $3E

            // NOTE: If we have turned on SRCHK function, all commands from this point on MUST include a Checksum!
            // NOTE: After we submit this comment we should see the Gateway respond with a sting as follows
            //    86uuuu0032pp00cc<cr><lf>
            //       pp = Command we specified
            //       cc = Checksum
            //       uu = the C-Bus unit address number of the Serial Gateway
            //
            
            
            // Ask for all current light/group level reports
            Logger.DebugFormat("{0}: Querying for Light Level Updates", this.DeviceName);
            getGroupLevelReport();
            
            // Finally, lets indicate that the driver is ready for use.
            IsReady = true;
        }


        void _serial_ConnectionLost(object sender, EventArgs e)
        {
            Logger.WarningFormat("{0}: Lost Serial Connection, the driver will keep trying to reconnect.", this.DeviceName);

            // We lost communications, so the driver is nolonger in ready status.
            IsReady = false;
        }


        void _timer_Elapsed(object sender, ElapsedEventArgs e)
        {
            try
            {
                // Every few minutes we can ask the system to update the Level report incase we missed a status update.
                Logger.DebugFormat("{0}: Querying for Light Level Updates", this.DeviceName);
                getGroupLevelReport();

            }
            catch
            {
            }
            finally
            {
                _timer.Start();
            }
        }


        #endregion

        #region Driver Payload Processing

        /// <summary>
        /// Review the data coming back from the Matrix
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void _serial_ResponseReceived(object sender, ReceivedDelimitedStringEventArgs e)
        {
            try
            {
                string response = e.RawResponse;

                // Conver the HEX String to a Hex Byte Array
                byte[] payload = HexStringToByteArray(e.RawResponse);

                
                Logger.DebugFormat("{0}: Recived Data {1}   {2}", this.DeviceName, response, payload.ToHexString(" "));
                
                
                // Level Status
                // If (Payload Code = Level from This Interface) or (Payload Code = Level from Elsewhere)
                // if (payload[1] = 0x07) or (payload[1] = 0x47) 
                //    length = Status Header - E0
                //    int length = (int)payload[0] - 192;
                //    Initial Block = (int)payload[3]
                //    int index = payload[3];

                // SH CO AP BS 
                // SH = E0 - F9
                // F9 07 38 2B AA AA AA AA AA AA 95 66 AA AA 5A 5A AA AA AA AA AA AA A5 A5 AA AA 04
                // F7 07 38 36 AA AA AA AA AA AA AA AA AA AA AA AA AA AA AA AA AA AA AA AA 4C
                // F9 07 38 40 00 00 AA AA AA AA 5A 5A 55 55 AA AA AA AA AA AA AA AA 00 00 55 95 48 
                // F9 07 38 4B 00 00 00 00 00 00 00 00 00 00 AA AA AA AA AA AA AA AA AA AA 00 00 D9 
                // F7 07 38 56 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 74 
                //
 
                // Lighting Application Report
                if ((payload[1] == 0x38) || (payload[2] == 0x38))
                {
                    if ((payload[1] == 0x07) || (payload[1] == 0x47))
                    {
                        // Level Status Message. The Status Header will be E0 + number of Bytes to follow

                        // : Message Format
                        // "Status + Coding + Application ID + Block Start + Data Bytes ... 
                        //
                        // : Example Message
                        //  00 01 02 03 04 05 06 07 08 09 10 11 12 13 14 15 16 17 18 19 20 21 22 23 24 25
                        //              00 01 02 03 04 05 06 07 08 09 10 11 12 13 14 15 16 17 18 19 20 21 
                        //              00    01    02    03    04    05    06    07    08    09    10 
                        // "F9 07 38 2B AA AA AA AA AA AA 95 66 AA AA 5A 5A AA AA AA AA AA AA A5 A5 AA AA 04"
                        //
                        // : Expansion
                        //   F9 = The length of this message - 0xE0 (In this case 0x19 or 25 Bytes)
                        //   07 = The Coding Type where this is 0x07 or 0x47 to indicate Level Data
                        //   38 = The Application this message is sourced from - 0x38 is Clipsal Lighting
                        //   00 = The start index that this block of data is referencing
                        //   xx.= this is the status of the lights, each byte contains 4 pairs of bits 
                        //        with each pair representing one light, thus 4 lights reported in 1 byte
                        
                        int length = (int)payload[0] - 224;
                        length = length - 3;                        // The initial calculation also included the non Payload bytes and Checksum
                        length = (length / 2) - 1;                  // 2 Bytes per Group, so divide the length in 1/2 
                        int index = payload[3];
                        

                        // What have we here?
                        // We have a CAL Message, Lenth
                        Logger.DebugFormat("{0}: Processing Level Message Block {1} for {2} Groups", this.DeviceName, index, length);

                        // We now need to process each byte of data in the block to extract the actual status information
                        for (int i = 0; i <= length; i++)
                        {
                            // Level information for a single Group Address Variable is reported in 2 bytes (4 characters)
                            // The bytes are split into 4 bit nibbles, exch of which can only be $5, $6, $9, or $A (11,10,01, and 00)
                            
                            int dataByte = 4 + (i * 2);
                            int groupID = index + i;                         // The start index + Current Byte Pair
                            
                            
                            // Sample - 95 99     [10010101  10011001]
                            // Sample - AA AA     [10101010  10101010]
                            // Sample - 55 55     [01010101  01010101]
                            // Sample - 5A 5A     [01011010  01011010]

                            byte mask = 15;

                            uint nibble4 = payload[dataByte + 1];
                            nibble4 = nibble4 >> 4;                         // 10010101 >> 4 = 00001001 [$9] 01
                            
                            uint nibble3 = payload[dataByte + 1];
                            //nibble3 = nibble3 << 4;                         // 10010101 << 4 = 01010000
                            //nibble3 = nibble3 >> 4;                         // 01010000 >> 4 = 00000101 [$5] 01
                            nibble3 = nibble3 & mask;
                            

                            uint nibble2 = payload[dataByte];
                            nibble2 = nibble2 >> 4;                         // 10010101 >> 4 = 00001001 [$9] 01
                            
                            uint nibble1 = payload[dataByte];
                            //nibble1 = nibble1 << 4;                         // 10010101 << 4 = 01010000
                            //nibble1 = nibble1 >> 4;                         // 01010000 >> 4 = 00000101 [$5] 11
                            nibble1 = nibble1 & mask;


                            uint groupLevelCalculation = 0;
                            uint groupLevel = 0;

                            // Now that we have each of the Nibbles in place, each Nibble has a defined binary value
                            // we need to add these values in the 4 nibbles to create a byte representing the level of the group

                            // n4 + n3 + n2 + n1
                            // 5    A    5    A
                            // 11   00   11   00
                            // C0   30   0C   00

                            switch (nibble1)
                            {
                                case 0x05: 
                                    // 00 00 00 11
                                    groupLevel = 0x03;
                                    break;

                                case 0x06: 
                                    // 00 00 00 10
                                    groupLevel = 0x02;
                                    break;

                                case 0x09: 
                                    // 00 00 00 01
                                    groupLevel = 0x01;
                                    break;

                                case 0x0A: 
                                    // 00 00 00 00
                                    groupLevel = 0x00;
                                    break;
                            }


                            switch (nibble2)
                            {
                                case 0x05: 
                                    // 00 00 11 00
                                    groupLevelCalculation = 0x0C;
                                    break;

                                case 0x06: 
                                    // 00 00 10 00
                                    groupLevelCalculation = 0x08;
                                    break;

                                case 0x09: 
                                    // 00 00 01 00
                                    groupLevelCalculation = 0x04;
                                    break;

                                case 0x0A: 
                                    // 00 00 00 00
                                    groupLevelCalculation = 0x00;
                                    break;
                            }

                            groupLevel = groupLevel + groupLevelCalculation;


                            switch (nibble3)
                            {
                                case 0x05:
                                    // 00 11 00 00
                                    groupLevelCalculation = 0x30;
                                    break;

                                case 0x06:
                                    // 00 10 00 00
                                    groupLevelCalculation = 0x20;
                                    break;

                                case 0x09:
                                    // 00 01 00 00
                                    groupLevelCalculation = 0x10;
                                    break;

                                case 0x0A:
                                    // 00 00 00 00
                                    groupLevelCalculation = 0x00;
                                    break;
                            }

                            groupLevel = groupLevel + groupLevelCalculation;


                            switch (nibble4)
                            {
                                case 0x05:
                                    // 11 00 00 00
                                    groupLevelCalculation = 0xC0;
                                    break;

                                case 0x06:
                                    // 10 00 00 00
                                    groupLevelCalculation = 0x80;
                                    break;

                                case 0x09:
                                    // 01 00 00 00
                                    groupLevelCalculation = 0x40;
                                    break;

                                case 0x0A:
                                    // 00 00 00 00
                                    groupLevelCalculation = 0x00;
                                    break;
                            }

                            groupLevel = groupLevel + groupLevelCalculation;

                            
                            _lightPowerLevels[groupID] = (int)groupLevel;

                            Logger.DebugFormat("{0}: LMB Light Group {1} - {2} Level {3}", this.DeviceName, groupID, _lightNames[groupID], _lightPowerLevels[groupID]);

                            DevicePropertyChangeNotification("LightLevels", groupID, _lightPowerLevels[groupID]);

                        }
                    }
                    else
                    {
                        // CAL Data Recieved. The Status Header will be C0 + number of Bytes to follow

                        // : Message Format
                        //  "Status + Application ID + Block Start + Data Bytes ... + Checksum"
                        // 
                        // : Example Message
                        //  "D8 38 00 A8 AA AA A6 AA A9 6A AA 6A A5 AA 9A A9 A9 AA AA 68 A9 12 00 AA 02 1F"
                        //
                        // : Expansion
                        //   D8 = The length of this message - 0xC0 (In this case 0x18 or 24 Bytes)
                        //   38 = The Application this message is sourced from - 0x38 is Clipsal Lighting
                        //   00 = The start index that this block of data is referencing
                        //   xx.= this is the status of the lights, each byte contains 4 pairs of bits 
                        //        with each pair representing one light, thus 4 lights reported in 1 byte
                        //

                        int length = (int)payload[0] - 192;
                        length = length - 2;                        // The initial calculation also included the non Payload bytes
                        int index = payload[2];


                        // What have we here?
                        // We have a CAL Message, Lenth
                        Logger.DebugFormat("{0}: Processing CAL Message Block {1} for {2} Groups", this.DeviceName, index, length * 4);

                        // We now need to process each byte of data in the block to extract the actual status information
                        for (int i = 0; i < length; i++)
                        {
                            // Each Byte Contains 4 Pairs of BITs, with each pair representing the status of a light
                            // The BITs have the following Interpetation
                            //   00 - 0 - This Group Address does not exisit on the network
                            //   01 - 1 - This Group Address is ON
                            //   10 - 2 - This Group Address is OFF
                            //   11 - 3 - This Group Address is ERROR state
                            //Logger.Debug("{0}:", payload[i +1].ToHexString()


                            int groupID = index + (i * 4);                  // The start index + (block number * 4 groups per block)
                            byte bits = payload[i + 3];                     // The first Payload packet will start at index 3

                            //     128 64  32  16  8   4   2   1 Calculation        Variable  Result   On Off
                            //     1   0   0   1   0   1   1   1 (128+16+4+2+1)     bits      7 
                            //
                            //     0   0   0   0   0   0   1   1 (2+1)              MaskA     3 
                            //     0   0   0   0   0   0   1   1 (1)                groupA    3        1  2

                            byte maskA = 3;
                            int groupA = bits & maskA;
                            if (groupA == 1)
                                _lightPowerStates[groupID + 0] = true;
                            else
                                _lightPowerStates[groupID + 0] = false;

                            //     0   0   0   0   1   1   0   0 (8+4)              MaskB     12 
                            //     0   0   0   0   0   1   0   0 (4)                groupB    4        4  8

                            byte maskB = 12;
                            int groupB = bits & maskB;
                            if (groupB == 4)
                                _lightPowerStates[groupID + 1] = true;
                            else
                                _lightPowerStates[groupID + 1] = false;

                            //     0   0   1   1   0   0   0   0 (32+16)            MaskC     48 
                            //     0   0   0   1   0   0   0   0 (16)               groupC    16       16 32

                            byte maskC = 48;
                            int groupC = bits & maskC;
                            if (groupC == 16)
                                _lightPowerStates[groupID + 2] = true;
                            else
                                _lightPowerStates[groupID + 2] = false;

                            //     1   1   0   0   0   0   0   0 (128+64)           MaskD     192 
                            //     1   0   0   0   0   0   0   0 (128)              groupD    128      64 128

                            byte maskD = 192;
                            int groupD = bits & maskD;
                            if (groupD == 64)
                                _lightPowerStates[groupID + 3] = true;
                            else
                                _lightPowerStates[groupID + 3] = false;


                            //Logger.DebugFormat("{0}: CAL Light Group {1} - {2} Power {3}", this.DeviceName, groupID + 0, _lightNames[groupID + 0], _lightPowerStates[groupID + 0]);
                            //Logger.DebugFormat("{0}: CAL Light Group {1} - {2} Power {3}", this.DeviceName, groupID + 1, _lightNames[groupID + 1], _lightPowerStates[groupID + 1]);
                            //Logger.DebugFormat("{0}: CAL Light Group {1} - {2} Power {3}", this.DeviceName, groupID + 2, _lightNames[groupID + 2], _lightPowerStates[groupID + 2]);
                            //Logger.DebugFormat("{0}: CAL Light Group {1} - {2} Power {3}", this.DeviceName, groupID + 3, _lightNames[groupID + 3], _lightPowerStates[groupID + 3]);


                            DevicePropertyChangeNotification("LightOnOffs", groupID + 0, _lightPowerStates[groupID + 0]);
                            DevicePropertyChangeNotification("LightOnOffs", groupID + 1, _lightPowerStates[groupID + 1]);
                            DevicePropertyChangeNotification("LightOnOffs", groupID + 2, _lightPowerStates[groupID + 2]);
                            DevicePropertyChangeNotification("LightOnOffs", groupID + 3, _lightPowerStates[groupID + 3]);
                        }
                    }
                }
                

                if (payload[0] == 0x05)
                {
                    // Point to Multipoint Data Recieved. 

                    // : Message Format
                    //  "Status + Application ID + Block Start + Data Bytes ... + Checksum"
                    //  "Header + Serial No + Application ID + $00 + Command + groupID + [ + Command Value ] + Checksum
                    // 
                    // : Example Message
                    //  "05 15 38 00 02 24 6A 1E"
                    //  "05 02 38 00 12 3C 01 10 D2 CC 12 43 CC 66
                    //  "05 02 38 00 12 31 00 12 4A 7F 12 34 4C 12 25 FF 12 32 00 97"
                    //
                    //
                    // : Expansion
                    //   05 = Point to Multipoint Data Packet
                    //   xx = Serial No of the sender... (Not Needed)
                    //   38 = The Application this message is sourced from - 0x38 is Clipsal Lighting
                    //   00 = The start index that this block of data is referencing
                    //   xx.= this is the status of the lights, each byte contains 4 pairs of bits 
                    //        with each pair representing one light, thus 4 lights reported in 1 byte
                    //

                    int startByte = 4;                            // the Start byte index for this payload
                    int length = payload.Length;
                    length = length - 4;                          // we dont need the header information for the next process
 
                    // This message recived can contain multiple status reports so we need to loop trough the message in blocks
                    for (int i = 0; i < length; i++)
                    {
                        int groupID = startByte + (i * 3);                          // index  * 3 bytes per block
                        int command = payload[groupID + 0];                         // The Command Type Been Processes
                        int index   = payload[groupID + 1];                         // The Group ID we are listening to
                        int level   = payload[groupID + 2];                         // The Level which this Group is Been Set To

                        Logger.DebugFormat("{0}: Processing P2M Data Block {1} - {2} Power {3}", this.DeviceName, index, command, level);

                        //if ((payload[groupID + 3] == 0x79) || (payload[groupID + 2] == 0x01))
                        if ((payload[command] == 0x79) || (payload[command] == 0x01))
                        {
                            // The Message is a Switch type, so we will override the level to either 0 for off, or 100 for on
                            if (command == 0x79)
                            {
                                level = 100;
                                _lightPowerStates[index] = true;
                            }
                            else
                            {
                                level = 0;
                                _lightPowerStates[index] = false;
                            }


                            Logger.DebugFormat("{0}: P2M Light Group {1} - {2} Power {3}", this.DeviceName, index, _lightNames[index], level);
                        }
                        else
                        {
                            // The Message is a Level type
                            Logger.DebugFormat("{0}: P2M Light Group {1} - {2} Level {3}", this.DeviceName, index, _lightNames[index], level);
                            _lightPowerLevels[index] = level;
                        }

                        
                        
                        DevicePropertyChangeNotification("LightLevels", index, _lightPowerLevels[index]);
                        DevicePropertyChangeNotification("LightOnOffs", index, _lightPowerStates[index]);
                    
                    }    
                }
            }
            catch (Exception ex)
            {
                Logger.ErrorFormat("{0}: Error in the Respone Recived - {1}", this.DeviceName, ex.Message);
            }

        }



        /// <summary>
        ///  Send the command to the matrix
        /// </summary>
        /// <param name="data"></param>
        private void sendCommand(string data)
        {
            Logger.DebugFormat("{0}: Sending Command - {1}", this.DeviceName, data);

            _serial.Send(data + TX_TERMINATOR);
        }


        private void getGroupLevelReport()
        {
            // : Light Status Requests Format
            // 
            // Format:
            //  \ + Point to Mutlipoint Header + $FF + $00 + STATUS (Request + Application + groupID)+ Checksum + Alpha Lowercase Char + CR    
            //
            // Payload:
            //  Header Field
            //    $03 : Point - Point - Multipoint, Lowest priority class
            //    $05 : Point - Multipoint, Lowest priority class
            //    $06 : Point - Point, Lowest priority class
            //  Status Field
            //    $7A       : Status Report - Switched
            //    $73 + $07 : Status Report - Level
            //  Application Field
            //    $38 : Standard Lighting Application [56]
            //  Lighting Group
            //    $00 : Hex ID of the Light Group to Target ($00, $20, $40, $60, $80, $A0, $C0, $E0)
            //  Checksum:
            //    Sum all of the bytes
            //    Find the remainder with the sum is devided by 256
            //    Take the 2's complement (invert the bits and then add 1) of the remainder
            //
            // Sample:
            //  \ + $05 + $FF + $00 + $73 + $07 + $38 + $groupID + $checksum
            // 
          
            for (int i = 0; i < 8; i++)
            {
                int groupID = i * 32;                          // index  * 3 bytes per block
                Logger.InfoFormat("{0}: Querying Light Level Base {1}", this.DeviceName, groupID);

                byte[] status = new byte[8];
                status[0] = 0x05;                       // Header      : Point - Multipoint, Lowest priority class
                status[1] = 0xFF;
                status[2] = 0x00;
                status[3] = 0x73;                       // Status      : Request Type 73 Level
                status[4] = 0x07;                       //             :              07 Level
                status[5] = 0x38;                       // Application : Standard Lighting Application [56]
                status[6] = (byte)groupID;              //Group to report status on

                // Checksum:
                //     Sum all of the bytes
                status[7] = (byte)(status[0] + status[1] + status[2] + status[3] + status[4] + status[5] + status[6]);
                //     Find the remainder with the sum is devided by 256
                status[7] = (byte)(status[7] % 256);
                //     Take the 2's complement (invert the bits and then add 1) of the remainder
                status[7] = (byte)(~status[7] + 1);

                sendCommand((string)"\\" + status.ToHexString());
            }
            
        }


        #endregion



        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        // Misc Device Functions
        //
        // The following method will allow us to directly access the control interface on the matrix
        //
        // We will create the following Functions:
        //
        //  Public Properties
        //      public ScriptArray ZoneNames
        //  Public Methods
        //      public void SendRawCommand(ScriptString rawCommand)
        //

        #region Misc Device Functions


        /// <summary>
        /// Expose a Property Array of the current zone names on the Matrix
        /// </summary>
        [ScriptObjectPropertyAttribute("Light Names", "Gets the name of all zones.", "The {NAME} zone name for zone #{INDEX|0}.", null)]
        public IScriptArray LightNames
        {
            get
            {
                // In the 1st constructor parameter, _zoneNames is of type string[].
                // In the 2nd constructor parameter, 1 is the index of the 1st element in the create ScriptArray.
                return new ScriptArrayMarshalByValue(_lightNames, 0);
            }
        }




        /// <summary>
        /// Expose a method for Raw Commmand passtrough so that the user has an ability to add additional functionality which is not enabled within the driver.
        /// There is no additional parsing of the input string passed.
        /// </summary>
        /// <param name="rawCommand">String to pass directly to the device</param>
        [ScriptObjectMethod("Raw Command", "Send a raw command to the projector")]
        [ScriptObjectMethodParameter("rawCommand", "The command to send")]
        public void SendRawCommand(ScriptString rawCommand)
        {
            sendCommand(rawCommand);
        }



        #endregion


        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        // Light Power
        //
        // The following methods will allow us to manage the power on the physical matrix zones
        //
        // We will create the following Functions:
        //
        //  Private Methods
        //      private void setZoneVolume(int zoneNumber, int volume)
        //  
        //  Public Properties
        //      public ScriptArray ZonePowerStates
        // 
        //  Public Methods
        //      public void SetZonePower(ScriptNumber zoneNumber, ScriptBoolean power)
        //      public void TurnZoneOff(ScriptNumber zoneNumber)
        //      public void TurnZoneOn(ScriptNumber zoneNumber)
        //      public void ToggleZonePower(ScriptNumber zoneNumber)
        //      public void TurnAllZonesOff()
        //      public void TurnAllZonesOn()
        //

        #region Light Power

        /// <summary>
        /// Private Function to actaully instruct the Matrix to change the Power on a Zone 
        /// </summary>
        /// <param name="zoneNumber">The Zone Number which we will be changing the Volume on</param>
        /// <param name="volume">The new Power setting we will be applying (On or Off)</param>
        private void setLightPower(int groupID, bool power)
        {
            // First, ensure that a valid Zone number is provided based on the number of actually configured zones on the matrix

            if ((groupID < 0) || (groupID > _configuredZoneCount))
                throw new ArgumentOutOfRangeException("groupID", "The Light Group number must be between 0 and " + _configuredZoneCount);

            // Convert the Power setting from Boolean On/Off|True/False to a string representation
            byte powerCommand;
            if (power)
                powerCommand = 0x79;
            else
                powerCommand = 0x01;


            // Record to the Log that we are about to action a Power Change request
            Logger.InfoFormat("{0}: Light Group {1} Set Power to {2}", this.DeviceName, groupID, power);

            // Format the Devices Command String with the paramaters zoneNumber and Power, and send the command.

            // : Light Power On Command Format
            // 
            // Format:
            //  \ + Header + Application + $00 + Command + groupID + Checksum
            //
            // Payload:
            //  Header Field
            //    $03 : Point - Point - Multipoint, Lowest priority class
            //    $05 : Point - Multipoint, Lowest priority class
            //    $06 : Point - Point, Lowest priority class
            //  Application Field
            //    $38 : Standard Lighting Application [56]
            //  Command Field
            //    $01 : Light Channel Off
            //    $79 : Light Channel On
            //  Lighting Group
            //    $00 : Hex ID of the Light Group to Target
            //  Checksum:
            //    Sum all of the bytes
            //    Find the remainder with the sum is devided by 256
            //    Take the 2's complement (invert the bits and then add 1) of the remainder
            //
            // Sample:
            //  \ + $05 + $38 + $00 + $79|$01 + $groupID + $checksum
            // 
            // Examples
            //  Group 37 OFF : \05 38 00 01 25 9D .
            //  Group 37 ON  : \05 38 00 79 25 25 .


            byte[] data = new byte[6];
            data[0] = 0x05;                     // Header:      Point - Multipoint, Lowest priority class
            data[1] = 0x38;                     // Application: Standard Lighting Application [56]
            data[2] = 0x00;                     // 00
            data[3] = powerCommand;             // Command:     $01 : Light Channel Off | $79 : Light Channel On
            data[4] = (byte)(groupID);          // groupID:     Hex Number of the Group to Target


            // Checksum:
            //     Sum all of the bytes
            data[5] = (byte)(data[0] + data[1] + data[2] + data[3] + data[4]);
            //     Find the remainder with the sum is devided by 256
            data[5] = (byte)(data[5] % 256);
            //     Take the 2's complement (invert the bits and then add 1) of the remainder
            data[5] = (byte)(~data[5] + 1);


            // Issue Command
            sendCommand("\\"+data.ToHexString());
        }



        /// <summary>
        /// Expose a Property Array of current Zone Power Status. All Power Settings are exposed as True/False
        /// </summary>
        [ScriptObjectPropertyAttribute("Light Group Power", "Gets or sets the current power setting for the Light Group. The setting is On/Off or True/False.", "the {NAME} power for Light Group #{INDEX|0}", "Set {NAME} Light Group #{INDEX|0} power to #{VALUE|false|On|Off}", typeof(ScriptBoolean), 0, MaxZoneCount, "LightNames")]
        [SupportsDriverPropertyBinding("Light Group Power State Changed", "Occurs when the current power setting for the zone changes.")]
        public IScriptArray LightOnOffs
        {
            get
            {
                // If you have an bool[] array dedicated to power states then you could do this:
                return new ScriptArrayMarshalByReference(_lightPowerStates, new ScriptArraySetBooleanCallback(setLightPower), 0);
            }
        }




        /// <summary>
        /// Exposes a method to set the Power on a specified zone in the Matrix
        /// </summary>
        /// <param name="zoneNumber">The Zone Number which we will be Changing the Power on</param>
        /// <param name="power">The new Power Setting we will be applying (True/False)</param>
        [ScriptObjectMethodAttribute("Set Light Group Power", "Change the Power Status for a Light Group.", "Set the Power for {NAME} Light Group {PARAM|0|1} to {PARAM|1|false|on|off}.")]
        [ScriptObjectMethodParameterAttribute("groupID", "The zone number.", 0, MaxZoneCount, "LightNames")]
        [ScriptObjectMethodParameter("Power", "Set the Power Status to ON or OFF", new string[] { "On", "Off" })]
        public void SetLightPower(ScriptNumber groupID, ScriptBoolean power)
        {
            setLightPower((int)groupID, (bool)power);
        }




        /// <summary>
        /// Exposes a method to Power Off a specified zone in the Matrix
        /// </summary>
        /// <param name="zoneNumber">The Zone Number we changing the power on</param>
        [ScriptObjectMethodAttribute("Turn Off Light", "Turn Off the Power for a Light Group.", "Turn Off the Power for {NAME} Light Group {PARAM|0|1}")]
        [ScriptObjectMethodParameterAttribute("groupID", "The light group number.", 0, MaxZoneCount, "LightNames")]
        public void TurnOffLight(ScriptNumber groupID)
        {
            // First, ensure that a valid Zone number is provided based on the number of actually configured zones on the matrix
            if ((groupID < 0) || (groupID > _configuredZoneCount))
                throw new ArgumentOutOfRangeException("groupID", "The Light Group number must be between 0 and " + _configuredZoneCount);

            // Record to the Log that we are about to action a Power Change request
            Logger.InfoFormat("{0}: Light Group {1} Power off", this.DeviceName, (int)groupID);

            // Format the Devices Command String with the paramaters zoneNumber and Power, and send the command.

            // : Light Power On Command Format
            // 
            // Format:
            //  \ + Point to Multipoint Header + Application + $00 + SAL DATA (Command + groupID) + Checksum + Alpha Lowercase Char + CR
            //
            // Payload:
            //  Header Field
            //    $03 : Point - Point - Multipoint, Lowest priority class
            //    $05 : Point - Multipoint, Lowest priority class
            //    $06 : Point - Point, Lowest priority class
            //  Application Field
            //    $38 : Standard Lighting Application [56]
            //  Command Field
            //    $01 : Light Channel Off
            //    $79 : Light Channel On
            //  Lighting Group
            //    $00 : Hex ID of the Light Group to Target
            //  Checksum:
            //    Sum all of the bytes
            //    Find the remainder with the sum is devided by 256
            //    Take the 2's complement (invert the bits and then add 1) of the remainder
            //
            // Sample:
            //  \ + $05 + $38 + $00 + $79|$01 + $groupID + $checksum
            // 
            // Examples
            //  Group 37 OFF : \05 38 00 01 25 9D .
            //  Group 37 ON  : \05 38 00 79 25 25 .


            byte[] data = new byte[6];
            data[0] = 0x05;                     // Header:      Point - Multipoint, Lowest priority class
            data[1] = 0x38;                     // Application: Standard Lighting Application [56]
            data[2] = 0x00;                     // 00
            data[3] = 0x01;                     // Command:     $01 : Light Channel Off
            data[4] = (byte)(groupID);          // groupID:     Hex Number of the Group to Target


            // Checksum:
            //     Sum all of the bytes
            data[5] = (byte)(data[0] + data[1] + data[2] + data[3] + data[4]);
            //     Find the remainder with the sum is devided by 256
            data[5] = (byte)(data[5] % 256);
            //     Take the 2's complement (invert the bits and then add 1) of the remainder
            data[5] = (byte)(~data[5] + 1);


            // Issue Command
            sendCommand("\\" + data.ToHexString());
        }





        /// <summary>
        /// Exposes a method to Power On a specified zone in the Matrix
        /// </summary>
        /// <param name="zoneNumber">The Zone Number we changing the power on</param>
        [ScriptObjectMethodAttribute("Turn On Light", "Turn On the Power for a Light Group.", "Turn On the Power for {NAME} Light Group {PARAM|0|1}")]
        [ScriptObjectMethodParameterAttribute("groupID", "The light group number.", 0, MaxZoneCount, "LightNames")]
        public void TurnOnLight(ScriptNumber groupID)
        {
            // First, ensure that a valid Zone number is provided based on the number of actually configured zones on the matrix
            if ((groupID < 0) || (groupID > _configuredZoneCount))
                throw new ArgumentOutOfRangeException("groupID", "The Light Group number must be between 0 and " + _configuredZoneCount);

            // Record to the Log that we are about to action a Power Change request
            Logger.InfoFormat("{0}: Light Group {1} Power on", this.DeviceName, (int)groupID);

            // Format the Devices Command String with the paramaters zoneNumber and Power, and send the command.

            // : Light Power On Command Format
            // 
            // Format:
            //  \ + Point to Multipoint Header + Application + $00 + SAL DATA (Command + groupID) + Checksum + Alpha Lowercase Char + CR
            //
            // Payload:
            //  Header Field
            //    $03 : Point - Point - Multipoint, Lowest priority class
            //    $05 : Point - Multipoint, Lowest priority class
            //    $06 : Point - Point, Lowest priority class
            //  Application Field
            //    $38 : Standard Lighting Application [56]
            //  Command Field
            //    $01 : Light Channel Off
            //    $79 : Light Channel On
            //  Lighting Group
            //    $00 : Hex ID of the Light Group to Target
            //  Checksum:
            //    Sum all of the bytes
            //    Find the remainder with the sum is devided by 256
            //    Take the 2's complement (invert the bits and then add 1) of the remainder
            //
            // Sample:
            //  \ + $05 + $38 + $00 + $79|$01 + $groupID + $checksum
            // 
            // Examples
            //  Group 37 OFF : \05 38 00 01 25 9D .
            //  Group 37 ON  : \05 38 00 79 25 25 .


            byte[] data = new byte[6];
            data[0] = 0x05;                     // Header:      Point - Multipoint, Lowest priority class
            data[1] = 0x38;                     // Application: Standard Lighting Application [56]
            data[2] = 0x00;                     // 00
            data[3] = 0x79;                     // Command:     $79 : Light Channel On
            data[4] = (byte)(groupID);          // groupID:     Hex Number of the Group to Target


            // Checksum:
            //     Sum all of the bytes
            data[5] = (byte)(data[0] + data[1] + data[2] + data[3] + data[4]);
            //     Find the remainder with the sum is devided by 256
            data[5] = (byte)(data[5] % 256);
            //     Take the 2's complement (invert the bits and then add 1) of the remainder
            data[5] = (byte)(~data[5] + 1);


            // Issue Command
            sendCommand("\\" + data.ToHexString());
        }


        /// <summary>
        /// Exposes a method to Power On a specified zone in the Matrix
        /// </summary>
        /// <param name="zoneNumber">The Zone Number we changing the power on</param>
        [ScriptObjectMethodAttribute("Toggle Light Power", "Toggle the Power for a Light Group.", "Toggle the Power for {NAME} Light Group {PARAM|0|1}")]
        [ScriptObjectMethodParameterAttribute("groupID", "The light group number.", 0, MaxZoneCount, "LightNames")]
        public void ToggleLightPower(ScriptNumber groupID)
        {
            // First, ensure that a valid Zone number is provided based on the number of actually configured zones on the matrix
            if ((groupID < 0) || (groupID > _configuredZoneCount))
                throw new ArgumentOutOfRangeException("groupID", "The Light Group number must be between 0 and " + _configuredZoneCount);


            if (_lightPowerStates[(int)groupID])
            {
                Logger.InfoFormat("{0}: Toggle Light Group {1} Power Off", this.DeviceName, (int)groupID);
                TurnOffLight(groupID);
            }
            else
            {
                Logger.InfoFormat("{0}: Toggle Light Group {1} Power On", this.DeviceName, (int)groupID);
                TurnOnLight(groupID);
            }

        }


        [ScriptObjectMethod("Turn On Light by Name", "Turns on the specified node.")]//, "Turn on {NAME} light {PARAM|0|change me}.")]
        [ScriptObjectMethodParameter("Name", "The name of the light.")]
        public void TurnOnLightByName(ScriptString name)
        {
            ScriptNumber id = findDeviceName((string)name);
            TurnOnLight(id);
        }


        [ScriptObjectMethod("Turn Off Light by Name", "Turns off the specified node.")]//, "Turn off {NAME} light #{PARAM|0|change me}.")]
        [ScriptObjectMethodParameter("Name", "The name of the light.")]
        public void TurnOffLightByName(ScriptString name)
        {
            ScriptNumber id = findDeviceName((string)name);
            TurnOffLight(id);
        }



        /// <summary>
        /// Exposes a method to Power Off all the zones on the Matrix
        /// </summary>
        [ScriptObjectMethodAttribute("Turn Off All Lights", "Power Down all Light Groups.", "Power Off all Light Groups")]
        public void TurnOffAllLights()
        {
            // Loop trought each of the zones on the device, and call the Method to power it off

            for (int zone = 0; zone < _configuredZoneCount; zone++)
            {
                TurnOffLight(new ScriptNumber(zone));
            }
        }




        /// <summary>
        /// Exposes a method to Power On all the zones on the Matrix
        /// </summary>
        [ScriptObjectMethodAttribute("Turn On All Lights", "Power Up all Light Groups.", "Power On all Light Groups")]
        public void TurnOnAllLights()
        {
            // Loop trought each of the zones on the device, and call the Method to power it on

            for (int zone = 0; zone < _configuredZoneCount; zone++)
            {
                TurnOnLight(new ScriptNumber(zone));
            }
        }


        #endregion

        


        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        // Light Level
        //
        // The following methods will allow us to manage the light level on the C-Bus Network
        //
        // We will create the following Functions:
        //
        //  Private Methods
        //      private void setZoneVolume(int zoneNumber, int volume)
        //  
        //  Public Properties
        //      public ScriptArray LightLevels
        // 
        //  Public Methods
        //      public void SetLightLevel(ScriptNumber zoneNumber, ScriptNumber volume)
        //      public void IncrementZoneVolume(ScriptNumber zoneNumber)
        //      public void DecrementZoneVolume(ScriptNumber zoneNumber)
        //

        #region Light Level

        /// <summary>
        /// Private Function to actaully instruct the Matrix to change the Level on a Light Group ID 
        /// </summary>
        /// <param name="gorupID">The Group ID which we will be changing the Level on</param>
        /// <param name="level">The new Light Level we will be applying (Range is 0 to 100)</param>
        private void setLightLevel(int groupID, int level)
        {
            // First, ensure that a valid Zone number is provided based on the number of actually configured zones on the matrix

            if ((groupID < 0) || (groupID > _configuredZoneCount))
                throw new ArgumentOutOfRangeException("GroupID", "The Group ID of the Lights must be between 0 and " + _configuredZoneCount);

            // Next, ensure that the volume level requested is within the range required of 0 to 100, round back if required.
            if (level < 0) level = 0;
            if (level > 100) level = 255;


            // Range on the C-Bus for Light Level is 0 to 255, which we will be exposing to the users in the scale of 0 to 100
            // to achieve this, we will do a simple match conversion when writing and reading the volume level from the device
            // as we store and retrive the base 100 result. Note that we will alway use Absoute numbers to communictate with the device
            //
            // Example - 
            //  If the user sets level to 100, we calculate the true level to be 255 
            //  If the user sets level to 50, we calculate the true volume to be 127
            //
            // Forumla -  
            //  groupLevel = level * 0.255

            byte groupLevel = (byte)Math.Abs(level * 2.55);

            // Record to the Log that we are about to action a Volume Change request
            Logger.InfoFormat("{0}: Group ID {1} Level set to {2}", this.DeviceName, groupID, groupLevel);


            // : Light Level Command Format
            // 
            // Format:
            //  \ + Point to Multipoint Header + Application + $00 + SAL DATA (Command + groupID + Level) + Checksum + Alpha Lowercase Char + CR
            //
            // Payload:
            //  Header Field
            //    $03 : Point - Point - Multipoint, Lowest priority class
            //    $05 : Point - Multipoint, Lowest priority class
            //    $06 : Point - Point, Lowest priority class
            //  Application Field
            //    $38 : Standard Lighting Application [56]
            //  Command Field
            //    $02 : 0 Second Ramp Rate (Instant)
            //    $0A : 4 Second Ramp Rate
            //    $12 : 8 Second Ramp Rate
            //    $1A : 12 Second Ramp Rate
            //    $22 : 20 Second Ramp Rate
            //    $2A : 30 Second Ramp Rate
            //    $32 : 40 Second Ramp Rate
            //    $3A : 60 Second Ramp Rate
            //    $42 : 90 Second Ramp Rate
            //    $4A : 120 Second Ramp Rate
            //    $52 : 180 Second Ramp Rate
            //    $5A : 5 Minute Ramp Rate
            //    $62 : 7 Minute Ramp Rate
            //    $6A : 10 Minute Ramp Rate
            //    $72 : 15 Minute Ramp Rate
            //    $7A : 17 Minute Ramp Rate
            //    $09 : Stop Ramp
            //  Lighting Group
            //    $00 : Hex ID of the Light Group to Target
            //  Checksum:
            //    Sum all of the bytes
            //    Find the remainder with the sum is devided by 256
            //    Take the 2's complement (invert the bits and then add 1) of the remainder
            //
            // Sample:
            //  \ + $05 + $38 + $00 + $02 + $groupID + $level + $checksum
            // 
            // Examples
            //  Group 52 10% : \05 38 00 02 34 19 74 .
            //  Group 52 30% : \05 38 00 02 34 4C 41 .
            //  Group 52 50% : \05 38 00 02 34 7F 0E .

            byte rampRate = 0x02;

            byte[] data = new byte[7];
            data[0] = 0x05;                     // Header:      Point - Multipoint, Lowest priority class
            data[1] = 0x38;                     // Application: Standard Lighting Application [56]
            data[2] = 0x00;                     // 00
            data[3] = rampRate;                 // Command:     Set to Instant Ramp Rate
            data[4] = (byte)(groupID);          // groupID:     Hex Number of the Group to Target
            data[5] = groupLevel;

            // Checksum:
            //     Sum all of the bytes
            data[6] = (byte)(data[0] + data[1] + data[2] + data[3] + data[4] + data[5]);
            //     Find the remainder with the sum is devided by 256
            data[6] = (byte)(data[6] % 256);
            //     Take the 2's complement (invert the bits and then add 1) of the remainder
            data[6] = (byte)(~data[6] + 1);


            // Issue Command
            sendCommand("\\" + data.ToHexString());
        }

        private void setLightLevel(int groupID, int level, byte rampRate)
        {
            // First, ensure that a valid Zone number is provided based on the number of actually configured zones on the matrix

            if ((groupID < 0) || (groupID > _configuredZoneCount))
                throw new ArgumentOutOfRangeException("GroupID", "The Group ID of the Lights must be between 0 and " + _configuredZoneCount);

            // Next, ensure that the volume level requested is within the range required of 0 to 100, round back if required.
            if (level < 0) level = 0;
            if (level > 100) level = 255;


            // Range on the C-Bus for Light Level is 0 to 255, which we will be exposing to the users in the scale of 0 to 100
            // to achieve this, we will do a simple match conversion when writing and reading the volume level from the device
            // as we store and retrive the base 100 result. Note that we will alway use Absoute numbers to communictate with the device
            //
            // Example - 
            //  If the user sets level to 100, we calculate the true level to be 255 
            //  If the user sets level to 50, we calculate the true volume to be 127
            //
            // Forumla -  
            //  groupLevel = level * 0.255

            byte groupLevel = (byte)Math.Abs(level * 2.55);

            // Record to the Log that we are about to action a Volume Change request
            Logger.InfoFormat("{0}: Group ID {1} Level set to {2}", this.DeviceName, groupID, groupLevel);


            // : Light Level Command Format
            // 
            // Format:
            //  \ + Point to Multipoint Header + Application + $00 + SAL DATA (Command + groupID + Level) + Checksum + Alpha Lowercase Char + CR
            //
            // Payload:
            //  Header Field
            //    $03 : Point - Point - Multipoint, Lowest priority class
            //    $05 : Point - Multipoint, Lowest priority class
            //    $06 : Point - Point, Lowest priority class
            //  Application Field
            //    $38 : Standard Lighting Application [56]
            //  Command Field
            //    $02 : 0 Second Ramp Rate (Instant)
            //    $0A : 4 Second Ramp Rate
            //    $12 : 8 Second Ramp Rate
            //    $1A : 12 Second Ramp Rate
            //    $22 : 20 Second Ramp Rate
            //    $2A : 30 Second Ramp Rate
            //    $32 : 40 Second Ramp Rate
            //    $3A : 60 Second Ramp Rate
            //    $42 : 90 Second Ramp Rate
            //    $4A : 120 Second Ramp Rate
            //    $52 : 180 Second Ramp Rate
            //    $5A : 5 Minute Ramp Rate
            //    $62 : 7 Minute Ramp Rate
            //    $6A : 10 Minute Ramp Rate
            //    $72 : 15 Minute Ramp Rate
            //    $7A : 17 Minute Ramp Rate
            //    $09 : Stop Ramp
            //  Lighting Group
            //    $00 : Hex ID of the Light Group to Target
            //  Checksum:
            //    Sum all of the bytes
            //    Find the remainder with the sum is devided by 256
            //    Take the 2's complement (invert the bits and then add 1) of the remainder
            //
            // Sample:
            //  \ + $05 + $38 + $00 + $02 + $groupID + $level + $checksum
            // 
            // Examples
            //  Group 52 10% : \05 38 00 02 34 19 74 .
            //  Group 52 30% : \05 38 00 02 34 4C 41 .
            //  Group 52 50% : \05 38 00 02 34 7F 0E .

            

            byte[] data = new byte[7];
            data[0] = 0x05;                     // Header:      Point - Multipoint, Lowest priority class
            data[1] = 0x38;                     // Application: Standard Lighting Application [56]
            data[2] = 0x00;                     // 00
            data[3] = rampRate;                 // Command:     Set to Instant Ramp Rate
            data[4] = (byte)(groupID);          // groupID:     Hex Number of the Group to Target
            data[5] = groupLevel;

            // Checksum:
            //     Sum all of the bytes
            data[6] = (byte)(data[0] + data[1] + data[2] + data[3] + data[4] + data[5]);
            //     Find the remainder with the sum is devided by 256
            data[6] = (byte)(data[6] % 256);
            //     Take the 2's complement (invert the bits and then add 1) of the remainder
            data[6] = (byte)(~data[6] + 1);


            // Issue Command
            sendCommand("\\" + data.ToHexString());
        }


        /// <summary>
        /// Expose a Property Array of current Zone Volume Levels. All Volume levels are exposed in the Range 0 to 100
        /// </summary>
        [ScriptObjectPropertyAttribute("Light Levels", "Gets or sets the current level setting for the Light Group ID. The scale is 0 to 100.", 0, 100, "the {NAME} Level for Light Group ID #{INDEX|1}", "Set {NAME} Light Group ID #{INDEX|1} Level to {value|100}.", typeof(ScriptNumber), 1, MaxZoneCount, "LightNames")]
        [SupportsDriverPropertyBinding("Light Level Changed", "Occurs when the current level setting for the light group changes.")]
        public IScriptArray LightLevels
        {
            get
            {
                // If you have an int[] array dedicated to volumes then you could do this:
                return new ScriptArrayMarshalByReference(_lightPowerLevels, new ScriptArraySetInt32Callback(setLightLevel), 1);
            }
        }



        /// <summary>
        /// Exposes a method to set the volume on a specified zone in the Matrix
        /// </summary>
        /// <param name="zoneNumber">The Zone Number which we will be Changing the Volume on</param>
        /// <param name="volume">The new Volume Level we will be applying (Range is 0 to 100)</param>
        [ScriptObjectMethodAttribute("Set Light Level", "Change the Level for a Ligth Group.", "Set the level for {NAME} Light Group ID #{PARAM|0|1} to {PARAM|1|10}.")]
        [ScriptObjectMethodParameterAttribute("groupID", "The zone number.", 0, MaxZoneCount, "LightNames")]
        [ScriptObjectMethodParameterAttribute("Level", "Level (0=Dark, 100=Bright)", 0, 100)]
        public void SetLightLevel(ScriptNumber groupID, ScriptNumber level)
        {
            setLightLevel((int)groupID, (int)level);
        }


        /// <summary>
        /// Exposes a method to set the volume on a specified zone in the Matrix
        /// </summary>
        /// <param name="zoneNumber">The Zone Number which we will be Changing the Volume on</param>
        /// <param name="volume">The new Volume Level we will be applying (Range is 0 to 100)</param>
        [ScriptObjectMethodAttribute("Ramp Light Level", "Change the Level for a Ligth Group of a period of time.", "Set the level for {NAME} Light Group ID #{PARAM|0|1} to {PARAM|1|10} ramped over {PARAM|2|4 Sec}.")]
        [ScriptObjectMethodParameterAttribute("groupID", "The zone number.", 0, MaxZoneCount, "LightNames")]
        [ScriptObjectMethodParameterAttribute("Level", "Level (0=Dark, 100=Bright)", 0, 100)]
        [ScriptObjectMethodParameterAttribute("Ramp", "The duration elapsed to reach the requested Level",
            new string[] { 
                "0 Sec", "4 Sec", "8 Sec", "12 Sec", "20 Sec", "30 Sec", "40 Sec", "60 Sec", "90 Sec", "2 Min", "3 Min", "5 Min", "7 Min", "10 Min", "15 Min", "17 Min", "Stop Ramp" })
        ]
        public void RampLightLevel(ScriptNumber groupID, ScriptNumber level, ScriptString ramp)
        {


            if (! RampRates.Keys.Contains(ramp))
            {
                Logger.ErrorFormat("{0}: invalid ramp rate {1}", this.DeviceName, ramp);
                throw new Exception("Invalid Ramp Rate");
            }

            Logger.DebugFormat("{0} Setting ramp rate to {1}", this.DeviceName, ramp);

            byte rampRate = RampRates[ramp];
            
            setLightLevel((int)groupID, (int)level, (byte)rampRate);
        }



        [ScriptObjectMethod("Set Light Level by Name", "Sets the specified node's level to the specified percent.")]//, "Set {NAME} light #{PARAM|0|change me} to {PARAM|1|99}%.")]
        [ScriptObjectMethodParameter("Name", "The name of the light.")]
        [ScriptObjectMethodParameter("PercentOn", "The percent level to set the light to. Valid values: 0 to 99 where 0 is typically off and 99 is fully on.", 0, 100)]
        public void SetLightLevelByName(ScriptString name, ScriptNumber percentOn)
        {
            ScriptNumber id = findDeviceName((string)name);
            SetLightLevel(id, new ScriptNumber(0));
        }


        #endregion

        private ScriptNumber findDeviceName(string Name)
        {
            int DeviceID = 1;
            
            return new ScriptNumber(DeviceID);
        }


        
        #region Hex Managment
        /// <summary>
        /// Creates a byte array from the hexadecimal string. Each two characters are combined
        /// to create one byte. First two hexadecimal characters become first byte in returned array.
        /// Non-hexadecimal characters are ignored. 
        /// </summary>
        /// <param name="hexData">string to convert to byte array</param>
        /// <returns>byte array, in the same left-to-right order as the hexString</returns>
        public static byte[] HexStringToByteArray(string hexData)
        {
            try
            {
                hexData = hexData.Replace(" ", ""); // remove any spaces
                
                byte[] buffer = new byte[hexData.Length / 2];
                for (int i = 0; i < hexData.Length; i += 2)
                    buffer[i / 2] = (byte)Convert.ToByte(hexData.Substring(i, 2), 16);
                return buffer;
            }
            catch (Exception ex)
            {
                throw new ArgumentException("The hexData string must contain contiguous 2 character hexidecimal values.", ex);
            }
        }

        /// <summary>
        /// Creates a byte array from the hexadecimal string. Each two characters are combined
        /// to create one byte. First two hexadecimal characters become first byte in returned array.
        /// Non-hexadecimal characters are ignored. 
        /// </summary>
        /// <param name="hexString">string to convert to byte array</param>
        /// <returns>byte array, in the same left-to-right order as the hexString</returns>
        // public static byte[] GetBytes(string hexString, out int discarded)
        public static byte[] GetBytes(string hexString)
        {
            int discarded;

            discarded = 0;
            string newString = "";
            char c;

            // remove all none A-F, 0-9, characters
            for (int i = 0; i < hexString.Length; i++)
            {
                c = hexString[i];
                if (IsHexDigit(c))
                    newString += c;
                else
                    discarded++;
            }

            // if odd number of characters, discard last character
            if (newString.Length % 2 != 0)
            {
                discarded++;
                newString = newString.Substring(0, newString.Length - 1);
            }

            int byteLength = newString.Length / 2;
            byte[] bytes = new byte[byteLength];
            string hex;
            int j = 0;
            for (int i = 0; i < bytes.Length; i++)
            {
                hex = new String(new Char[] { newString[j], newString[j + 1] });
                bytes[i] = HexToByte(hex);
                j = j + 2;
            }
            return bytes;
        }


        /// <summary>
        /// Returns true is c is a hexadecimal digit (A-F, a-f, 0-9)
        /// </summary>
        /// <param name="c">Character to test</param>
        /// <returns>true if hex digit, false if not</returns>
        public static bool IsHexDigit(Char c)
        {
            int numChar;
            int numA = Convert.ToInt32('A');
            int num1 = Convert.ToInt32('0');
            c = Char.ToUpper(c);
            numChar = Convert.ToInt32(c);
            if (numChar >= numA && numChar < (numA + 6))
                return true;
            if (numChar >= num1 && numChar < (num1 + 10))
                return true;
            return false;
        }


        /// <summary>
        /// Converts 1 or 2 character string into equivalant byte value
        /// </summary>
        /// <param name="hex">1 or 2 character string</param>
        /// <returns>byte</returns>
        private static byte HexToByte(string hex)
        {
            if (hex.Length > 2 || hex.Length <= 0)
                throw new ArgumentException("hex must be 1 or 2 characters in length");
            byte newByte = byte.Parse(hex, System.Globalization.NumberStyles.HexNumber);
            return newByte;
        }

        #endregion


    }

}

