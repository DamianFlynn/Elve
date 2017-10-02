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
using System.Xml.Linq;

namespace BuiltInDrivers
{

	/// <summary>
	/// 
	/// Manuals:
	/// http://www.xantech.com/products/av_distribution/MRC2%20RS232%20Comm.pdf
	///
	/// Overview:
	/// The Driver has been written to manage both the standard and expanded modes of the Xantec MRC 88
	/// Matrix Amplifiers
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
		"Xantech MRC88 Driver",                                                                         // Display Name 
		"This driver supports control and monitoring of the Xantech MRC controllers" +                  // Description
		"in Single or expanded configuration. With passtrough for IR Command Macros. " +
		"The serial interface on the Xantech is 9600 Baud, and if the driver is used for " +
		"expanded mode (16 zones) then the Xantech should be initialized and cabled for " +
		"operation in this mode.",
		"Damian Flynn",                                                                                 // Author
		"Multi-Room Audio",                                                                             // Category
		"",                                                                                             // Subcategory
		"XantechMRC88",                                                                                 // Default instance name
		DriverCommunicationPort.Serial,                                                                 // Default Communications Metod
		DriverMultipleInstances.MultiplePerDriverService,                                               // Allow multiple instances
		0,                                                                                              // Major Version
		93,                                                                                             // Minor Version
		DriverReleaseStages.Test,                                                                       // Release Stage
		"Xantech",                                                                                      // Manufacturer Name
		"http://www.xantech.com",                                                                       // Manufacturer ULR
		"GuEaPVzL2ObOiKmGk9HY+yFkLqhZ8jym1LRTCBCQd9Vvsxo8mUkB9g=="                                      // Registration - for offical use
		)]


	public class Xantech_MRC88_Driver : Driver, IMultiroomAudioDriver, IMatrixSwitcherDriver
	{

		#region Variables

		// Define Communications Protocol Variables
		private SerialCommunication _serial;
		private string _serialPortName;                 // "COM1"
		private const string TX_TERMINATOR = "";
		private const string RX_TERMINATOR = "\r";

		// Define a timer object as we will be executing peroiodic status checks
		private System.Timers.Timer _timer;

		// Driver Variables

		// Hardware Physical Limits

		private const int MaxZoneCount = 16;        // Zone Number. Range is 1..8 (1..16 if expanded)
		private const int MaxSourceCount = 8;       // Source Input Number. Range is 1..8
		private const int MaxVolumeRange = 38;      // Volume Setting. Range is 0..38. 
		private const int MaxBaseRange = 14;        // Bass/Treble Setting. Range is 0..14. See BASS/TREBLE LEVEL table below.
		private const int MaxTrebleRange = 14;
		private const int MaxBalanceRange = 63;     // Balance Setting. Range is 0..63. See BALANCE LEVEL table below.
		private const int MaxSupportedMacros = 127; // Macro Number. Range is 0..127.



		// Zone Details
		private int _configuredZoneCount;

		private string[] _zoneNames = new string[MaxZoneCount];
		private string[] _zoneSourceNames = new string[MaxZoneCount];
		private int[] _zoneSources = new int[MaxZoneCount];
		private bool[] _zonePowerStates = new bool[MaxZoneCount];
		private bool[] _zoneMuteStates = new bool[MaxZoneCount];
		private bool[] _zoneLoudnessStates = new bool[MaxZoneCount];
		private bool[] _zoneSharedStates = new bool[MaxZoneCount];
		private bool[] _zoneDndStates = new bool[MaxZoneCount];
		private int[] _zonePartyModeStates = new int[MaxZoneCount];
		private int[] _zoneBackgroundStates = new int[MaxZoneCount];
		private int[] _zoneBalanceLevels = new int[MaxZoneCount];
		private int[] _zoneTrebleLevels = new int[MaxZoneCount];
		private int[] _zoneBassLevels = new int[MaxZoneCount];
		private int[] _zoneStartVolumeLevels = new int[MaxZoneCount];
		private int[] _zoneVolumeLevels = new int[MaxZoneCount];
		private DateTime[] _zoneLastUpdated = new DateTime[MaxZoneCount];
		private string[] _zoneTextLine1Messages = new string[MaxZoneCount];
		private string[] _zoneTextLine2Messages = new string[MaxZoneCount];

		// Source Details
		private string[] _sourceNames = new string[MaxSourceCount];

		// Audio Attenuation Levels in dB Details

		private string[] _volumeAttenuationLevels = new string[] { "-78.75", "-75.00", "-71.25", "-67.50", "-63.75", "-60.00", "-56.25", "-52.50", "-50.00", "-47.50", "-45.00", "-42.50", "-40.00", "-37.50", "-35.00", "-32.50", "-30.00", "-27.50", "-25.00", "-23.75", "-22.50", "-21.25", "-20.00", "-18.75", "-17.50", "-16.25", "-15.00", "-13.75", "-12.50", "-11.25", "-10.00", "-8.75", "-7.50", "-6.25", "-5.00", "-3.75", "-2.50", "-1.25", "0" };
		private string[] _baseAttenuationLevels = new string[] { "-14", "-12", "-10", "-8", "-6", "-4", "-2", "0", "2", "4", "6", "8", "10", "12", "14" };
		private string[] _trebleAttenuationLevels = new string[] { "-14", "-12", "-10", "-8", "-6", "-4", "-2", "0", "2", "4", "6", "8", "10", "12", "14" };
		private string[] _balanceAttenuationLevels = new string[] { "Mute", "-37.5", "-36.25", "-35", "-33.75", "-32.5", "-31.25", "-30", "-28.75", "-27.5", "-26.25", "-25", "-23.75", "-22.5", "-21.25", "-20", "-18.75", "-17.5", "-16.25", "-15", "-13.75", "-12.5", "-11.25", "-10", "-8.75", "-7.5", "-6.25", "-5", "-3.75", "-2.5", "-1.25", "0" };

		#endregion

		/// <summary>
		/// Initialization of the driver to default values
		/// </summary>
		public Xantech_MRC88_Driver()
		{
			// Initialize zone related vars            
			for (int i = 0; i < MaxZoneCount; i++)
			{
				_zoneNames[i] = "Zone " + (i + 1);
				_zoneSourceNames[i] = "Uninitialized";
				_zoneSources[i] = -1;
				_zonePowerStates[i] = false;
				_zoneMuteStates[i] = false;
				_zoneLoudnessStates[i] = false;
				_zoneSharedStates[i] = false;
				_zoneDndStates[i] = false;
				_zonePartyModeStates[i] = -1;
				_zoneBackgroundStates[i] = -1;
				_zoneBalanceLevels[i] = -1;
				_zoneTrebleLevels[i] = -1;
				_zoneBassLevels[i] = -1;
				_zoneStartVolumeLevels[i] = -1;
				_zoneVolumeLevels[i] = -1;
				_zoneTextLine1Messages[i] = "";
				_zoneTextLine2Messages[i] = "";
				_zoneLastUpdated[i] = DateTime.Now.AddMinutes(-5);  // Set the initial update to be 5 minutes ago, which should be enough to triger a forced refresh
			}

			// Initialize source related vars            
			for (int i = 0; i < MaxSourceCount; i++)
			{
				_sourceNames[i] = "Source " + (i + 1);
			}

		}



		#region DriverSettingAttribute
		// Provide setting for the driver to be presented in the Elve managment studio


		/// <summary>
		/// Communcation Port we will Utilize for Communications
		/// </summary>
		[DriverSettingAttribute("Serial Port Name", "The name of the serial port that the Xantech MRC88 is connected to. Ex. COM1", typeof(SerialPortDeviceSettingEditor), null, true)]
		public string SerialPortNameSetting
		{
			set { _serialPortName = value; }
		}


		/// <summary>
		/// Configuration Mode of our Matrix
		/// </summary>
		[DriverSettingAttribute("Matrix Configuration", "The configuration mode of the matrix", new string[] { "Single", "Expanded" }, null, true)]
		public string SerialExpandedModeConfiguration
		{
			set
			{
				// Matrix Mode - Will be one of "Single", "Expanded" 
				if (value == "Expanded")
					_configuredZoneCount = 16;
				else
					_configuredZoneCount = 8;

				Logger.DebugFormat("{0}: Congigured for {1} Zones", this.DeviceName, _configuredZoneCount);
			}
		}


		/// <summary>
		/// Assign a friendly name for each zone on the Matrix
		/// </summary>
		[DriverSettingArrayNamesAttribute("Custom Zone Names", "Enter the name of each Zone output from the Matrix.", typeof(ArrayItemsDriverSettingEditor), "ZoneNames", 1, MaxZoneCount, "", false)]
		public string CustomZoneNames
		{
			set
			{
				if (string.IsNullOrEmpty(value) == false)
				{
					XElement element = XElement.Parse(value);
                    foreach (XElement node in element.Elements("Item"))
                        _zoneNames[(int)node.Attribute("Index") - 1] = node.Attribute("Name").Value;
                }
			}
		}



		/// <summary>
		/// Assign a friendly name for each source on the Matrix
		/// </summary>
		[DriverSettingArrayNamesAttribute("Custom Source Names", "Enter the name of each Source connected to the Matrix", typeof(ArrayItemsDriverSettingEditor), "SourceNames", 1, MaxSourceCount, "", false)]
		public string CustomSourceNames
		{
			set
			{
				if (string.IsNullOrEmpty(value) == false)
				{
					XElement element = XElement.Parse(value);
                    foreach (XElement node in element.Elements("Item"))
                        _sourceNames[(int)node.Attribute("Index") - 1] = node.Attribute("Name").Value;
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
			_serial = new SerialCommunication(_serialPortName, 9600, System.IO.Ports.Parity.None, 8, System.IO.Ports.StopBits.One);
			_serial.CurrentEncoding = new ASCIIEncoding();
			_serial.Delimiter = RX_TERMINATOR; // this is the incoming message delimiter
			_serial.ReceivedDelimitedString += new EventHandler<ReceivedDelimitedStringEventArgs>(_serial_ResponseReceived);

			// Define the Logger Object for the Serial Device
			_serial.Logger = Logger;

			// Set up connection monitor.
			_serial.ConnectionMonitorTimeout = 60000;                // ensure we receive data at least once every minute so we know the connection is alive
			_serial.ConnectionMonitorTestRequest = "!ZA1+";          // transmit a simple dummy command every 60 seconds to ensure the unit is still connected.
			_serial.ConnectionEstablished += new EventHandler<EventArgs>(_serial_ConnectionEstablished);
			_serial.ConnectionLost += new EventHandler<EventArgs>(_serial_ConnectionLost);
			_serial.StartConnectionMonitor();                   // this will also attempt to open the serial connection


			// set timer to get the current device state
			_timer = new System.Timers.Timer();
			_timer.AutoReset = false;
			_timer.Elapsed += new ElapsedEventHandler(_timer_Elapsed);
			_timer.Interval = 20000;                           // Trigger the timer every 20 seconds, this checks the driver is still in ready state, and gets current status
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
			Logger.InfoFormat("{0}: Opened Serial Connection. (This does not guarantee that the hardware is physically connected to the serial port).", this.DeviceName);

			// Enable/Disable Auto Update of the Zone Information when Zone Activity is Detected.
			sendCommand("!ZA1+");

			// Loop trought each Zone and query for thier current status detail.
			for (int QueryZone = 1; QueryZone <= _configuredZoneCount; QueryZone++)
			{
				Logger.InfoFormat("{0}: Requesting Initial Status Data and Metadata for zone {1}.", this.DeviceName, QueryZone);
				sendCommand("?" + QueryZone + "ZS+");
			}

			// Lets check that we have processed a status update for each zone, and indicate that the driver is ready for use
			// Loop trought each Zone and query for thier current status detail.
			for (int QueryZone = 1; QueryZone <= _configuredZoneCount; QueryZone++)
			{
				// Check how much time has elapsed since the last time the zone was updated.
				TimeSpan elapsed = DateTime.Now.Subtract(_zoneLastUpdated[QueryZone - 1]);

				Logger.DebugFormat("{0}: Verifying the Device is ready - we processed zone {1}, {2} seconds ago", this.DeviceName, QueryZone, elapsed.TotalSeconds);
				// if less than 120 seconds has elapsed then we know that this zone has been recently updated.
				if (elapsed.TotalSeconds < 120)
				{
					IsReady = true;
				}
				else
				{
					Logger.DebugFormat("{0}: Driver is not ready - Zone {1} has not been updated in {2} seconds!", this.DeviceName, QueryZone, elapsed.TotalSeconds);
					IsReady = false;
				}
			}

		}


		void _serial_ConnectionLost(object sender, EventArgs e)
		{
			Logger.WarningFormat("{0}: Lost Serial Connection, the driver will keep trying to reconnect.", this.DeviceName);

			// We lost communications, so the driver is no longer in ready status.
			IsReady = false;
		}


		void _timer_Elapsed(object sender, ElapsedEventArgs e)
		{
			try
			{
				// Lets check that we have processed a status update for each zone, and indicate that the driver is ready for use
				// Loop trought each Zone and query for thier current status detail.
				int zonesReady = 1;
				for (int QueryZone = 1; QueryZone <= _configuredZoneCount; QueryZone++)
				{
					// Check how much time has elapsed since the last time the zone was updated.
					TimeSpan elapsed = DateTime.Now.Subtract(_zoneLastUpdated[QueryZone - 1]);

					// if less than 180 (3 Minutes) seconds has elapsed then we know that this zone has been recently updated.
					if (elapsed.TotalSeconds < 180)
					{
						IsReady = true;
						zonesReady++;
					}
					else
					{
						Logger.InfoFormat("{0}: Driver is not ready - Zone {1} has not been updated in {2} seconds!", this.DeviceName, QueryZone, elapsed.TotalSeconds);
						IsReady = false;
					}

				}

				if (zonesReady == _configuredZoneCount)
				{
					Logger.InfoFormat("{0}: Pending Ready - All Zones have been updated within the last 180 seconds!", this.DeviceName);
				}

				if ((zonesReady == _configuredZoneCount) && (IsReady == true))
				{
					Logger.InfoFormat("{0}: Driver Ready!", this.DeviceName);
				}


				// Loop trought each Zone and query for thier current status detail.
				for (int QueryZone = 1; QueryZone <= _configuredZoneCount; QueryZone++)
				{
					// Check how much time has elapsed since the last time the zone was updated.
					TimeSpan elapsed = DateTime.Now.Subtract(_zoneLastUpdated[QueryZone - 1]);

					// if more than 90 seconds have elapsed lets check that the information we have is accurate.
					if (elapsed.TotalSeconds > 90)
					{
						Logger.DebugFormat("{0}: Requesting Refresh Status Data for zone {1}, {2} seconds since last update.", this.DeviceName, QueryZone, elapsed.TotalSeconds);
						sendCommand("?" + QueryZone + "ZD+");
					}
				}
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

				switch ((char)response.ElementAt(0))
				{

					case 'O': // O - Ok Acknolwledgement
						Logger.DebugFormat("{0}: Recived OK Reply", this.DeviceName);
						break;

					case 'E': // E - Error Notice
						Logger.DebugFormat("{0}: Recived ERROR Reply", this.DeviceName);
						break;

					case '?': // ? - Query Reply
						// Sample:        
						// ?5PR0+
						Logger.DebugFormat("{0}: Recived QUERY Reply - {1}", this.DeviceName, response);
						break;

					case '#': // # - Status Reply
						Logger.DebugFormat("{0}: Recived STATUS Reply - {1}", this.DeviceName, response);
						ProcessResponseMessage(response);
						break;


					default:
						Logger.WarningFormat("{0}: Recived UNKNOWN Reply - {1}", this.DeviceName, response);
						return;
				}

			}
			catch (Exception ex)
			{
				Logger.ErrorFormat("{0}: Error in the Respone Recived - {1}", this.DeviceName, ex.Message);
			}

		}


		public void ProcessResponseMessage(string statusMessage)
		{
			/***************************************************************************
			* Status Message processor.
			* This device is capable of delivering 2 different status messages
			* 
			* ZoneStatus in the format:   #5ZS PR0 SS2 VO8 MU1 TR7 BS7 BA32 LS0 PS0+
			* ZoneMessage in the format:  #5ZM SS2 LT1 ""+
			*
			* The status messages are delimited by a space character, breaking the 
			* message into digestable content, for us to process easily
			***************************************************************************/

			/*
			* Block 1 on the message identifes the Zone this message represents along
			* with the type of message to be processed Status Block or Message Block
			*/

			int ZoneID = int.Parse(statusMessage.Substring(1, statusMessage.IndexOf("Z") - 1));
			string messageFormat = statusMessage.Substring(statusMessage.IndexOf("Z"), 2);

			/*
			* Next we will remove this first block and trailing + from the message block
			* before we being the full workload
			*/

			statusMessage = statusMessage.Substring(statusMessage.IndexOf(" ") + 1);
			statusMessage = statusMessage.Substring(0, statusMessage.Length - 1);


			/*
			 * Now we are ready to breakdown the blocks to extract the details reported to us.
			 */

			switch (messageFormat)
			{
				case "ZS":
					/// <summary>
					/// 
					/// ZS Payload
					///
					/// The message string will appear as: PR0 SS2 VO8 MU1 TR7 BS7 BA32 LS0 PS0
					///
					/// This information contained here is in blocks of sub messages
					///   PR0  => Power Report     [PR] with 0 for Off, 1 for On
					///   SS2  => Source connected [SS] followed by source number 1 to 8
					///   VO8  => Volume report    [VO] followed by the volume 0 to 30
					///   MU1  => Mute Report      [MU] with 0 for Off, 1 for On
					///   TR7  => Treble report    [TR] followed by the level 0 to 30
					///   BS7  => Bass report      [BS] followed by the level 0 to 30
					///   BA32 => Ballance report  [BA] followed by the level 0 to 30
					///   LS0  => Link Status      [LS] with 0 for Off, 1 for On
					///   PS0  => Paging Status    [PS] with 0 for Off, 1 for On
					///   
					/// To process these blocks, we will split each into array elements 
					/// this will be in the format of XXYY with XX = Setting and YY = Value
					/// 
					/// </summary>

					string logText = "";

					// We will break the message into blocks delimited by Spaces, and loop trought the list of blocks
					string[] statusBlocks = statusMessage.Split(new Char[] { ' ' });

					foreach (string statusDetail in statusBlocks)
					{
						// within each block we will have 2 characters to identify the payload, and the remainder should contain the value associated
						int statusValue = int.Parse(statusDetail.Substring(2));

						switch (statusDetail.Substring(0, 2))
						{
							case "PR":
								// Power Status Message
								//Logger.DebugFormat("{0}: Recived Zone {1} Status: Power is {2}", this.DeviceName, ZoneID, statusValue);
								logText = logText + "Power: " + statusValue + ". ";

								if (statusValue == 1)
									_zonePowerStates[ZoneID - 1] = true;
								else
									_zonePowerStates[ZoneID - 1] = false;

								DevicePropertyChangeNotification("ZonePowerStates", ZoneID, _zonePowerStates[ZoneID - 1]);
								break;

							case "MU":
								// Mute Status Message
								//Logger.DebugFormat("{0}: Recived Zone {1} Status: Mute is {2}", this.DeviceName, ZoneID, statusValue);
								logText = logText + "Mute: " + statusValue + ". ";

								if (statusValue == 1)
									_zoneMuteStates[ZoneID - 1] = true;
								else
									_zoneMuteStates[ZoneID - 1] = false;

								DevicePropertyChangeNotification("ZoneMuteStates", ZoneID, _zoneMuteStates[ZoneID - 1]);
								break;

							case "VO":
								// Volume Status Message
								//Logger.DebugFormat("{0}: Recived Zone {1} Status: Volume is {2}", this.DeviceName, ZoneID, statusValue);
								logText = logText + "Vol: " + statusValue + ". ";

								// Range on the Matrix for Volume is 0 to 38, which we will be exposing to the users in the scale of 0 to 100
								// to achieve this, we will do a simple match conversion when writing and reading the volume level from the device
								// as wer store and retrive the base 100 result. Note that we will alway use Absoute numbers to communictate with the device
								//
								// example - 
								//  If the device reports the volume to be 38, we calculate the presented volume to be 100
								//  If the device reports the volume to be 19, we calculate the presented volume to be 50
								//
								// Forumla =  volume = valuesupplied / 0.38

								int vol = (byte)Math.Abs(statusValue / 0.38);

								_zoneVolumeLevels[ZoneID - 1] = vol;

								// Notify the system that the Volume property's value may have changed. Note that you don't  
								// need to check if the _volume value actually changed or if it was just set the to the same  
								// value that it was previously set to. You can just invoke DevicePropertyChangeNotification 

								DevicePropertyChangeNotification("ZoneVolumes", ZoneID, _zoneVolumeLevels[ZoneID - 1]);
								break;

							case "BS":
								// Volume Status Message
								//Logger.DebugFormat("{0}: Recived Zone {1} Status: Bass is {2}", this.DeviceName, ZoneID, statusValue);
								logText = logText + "Bass: " + statusValue + ". ";

								// Range on the Matrix for Volume is 0 to 38, which we will be exposing to the users in the scale of 0 to 100
								// to achieve this, we will do a simple match conversion when writing and reading the volume level from the device
								// as wer store and retrive the base 100 result. Note that we will alway use Absoute numbers to communictate with the device
								//
								// example - 
								//  If the device reports the volume to be 38, we calculate the presented volume to be 100
								//  If the device reports the volume to be 19, we calculate the presented volume to be 50
								//
								// Forumla =  volume = valuesupplied / 0.14

								int bass = (byte)Math.Abs(statusValue / 0.14);

								_zoneBassLevels[ZoneID - 1] = bass;

								// Notify the system that the Volume property's value may have changed. Note that you don't  
								// need to check if the _volume value actually changed or if it was just set the to the same  
								// value that it was previously set to. You can just invoke DevicePropertyChangeNotification 

								DevicePropertyChangeNotification("ZoneBassLevels", ZoneID, _zoneBassLevels[ZoneID - 1]);
								break;


							case "TR":
								// Volume Status Message
								//Logger.DebugFormat("{0}: Recived Zone {1} Status: Treble is {2}", this.DeviceName, ZoneID, statusValue);
								logText = logText + "Treb: " + statusValue + ". ";

								// Range on the Matrix for Volume is 0 to 38, which we will be exposing to the users in the scale of 0 to 100
								// to achieve this, we will do a simple match conversion when writing and reading the volume level from the device
								// as wer store and retrive the base 100 result. Note that we will alway use Absoute numbers to communictate with the device
								//
								// example - 
								//  If the device reports the volume to be 38, we calculate the presented volume to be 100
								//  If the device reports the volume to be 19, we calculate the presented volume to be 50
								//
								// Forumla =  volume = valuesupplied / 0.14

								int treble = (byte)Math.Abs(statusValue / 0.14);

								_zoneTrebleLevels[ZoneID - 1] = treble;

								// Notify the system that the Volume property's value may have changed. Note that you don't  
								// need to check if the _volume value actually changed or if it was just set the to the same  
								// value that it was previously set to. You can just invoke DevicePropertyChangeNotification 

								DevicePropertyChangeNotification("ZoneTrebleLevels", ZoneID, _zoneTrebleLevels[ZoneID - 1]);
								break;


							case "BA":
								// Volume Status Message
								//Logger.DebugFormat("{0}: Recived Zone {1} Status: Balance is {2}", this.DeviceName, ZoneID, statusValue);
								logText = logText + "Bal: " + statusValue + ". ";

								// Range on the Matrix for Volume is 0 to 38, which we will be exposing to the users in the scale of 0 to 100
								// to achieve this, we will do a simple match conversion when writing and reading the volume level from the device
								// as wer store and retrive the base 100 result. Note that we will alway use Absoute numbers to communictate with the device
								//
								// example - 
								//  If the device reports the volume to be 38, we calculate the presented volume to be 100
								//  If the device reports the volume to be 19, we calculate the presented volume to be 50
								//
								// Forumla =  volume = valuesupplied / 0.64

								int balance = (byte)Math.Abs(statusValue / 0.64);

								_zoneBalanceLevels[ZoneID - 1] = balance;

								// Notify the system that the Volume property's value may have changed. Note that you don't  
								// need to check if the _volume value actually changed or if it was just set the to the same  
								// value that it was previously set to. You can just invoke DevicePropertyChangeNotification 

								DevicePropertyChangeNotification("ZoneBalanceLevels", ZoneID, _zoneBalanceLevels[ZoneID - 1]);
								break;


							case "SS":
								// Volume Status Message
								//Logger.DebugFormat("{0}: Recived Zone {1} Status: Source is {2}", this.DeviceName, ZoneID, statusValue);
								logText = logText + "Source: " + statusValue + ". ";

								_zoneSources[ZoneID - 1] = statusValue;
								_zoneSourceNames[ZoneID - 1] = _sourceNames[statusValue];

								// Notify the system that the Volume property's value may have changed. Note that you don't  
								// need to check if the _volume value actually changed or if it was just set the to the same  
								// value that it was previously set to. You can just invoke DevicePropertyChangeNotification 

								DevicePropertyChangeNotification("ZoneSources", ZoneID, _zoneSources[ZoneID - 1]);
                                DevicePropertyChangeNotification("ZoneSourceNames", ZoneID, _zoneSourceNames[ZoneID - 1]);
								break;

							default:
								// Any remaining unporcessed Messages
								//Logger.DebugFormat("{0}: Recived Zone {1} Status: Data Block {2} is {3}", this.DeviceName, ZoneID, statusDetail.Substring(0, 2), statusValue);
								logText = logText + statusDetail.Substring(0, 2) + ": " + statusValue + ". ";
								break;

						}

						// We will record for the timer that we just updated status information for this zone.
						_zoneLastUpdated[ZoneID - 1] = DateTime.Now;
					}

					Logger.DebugFormat("{0}: Recived Zone {1} {2}", this.DeviceName, ZoneID, logText);

					break;

				case "ZM":
					/// <summary>
					/// 
					/// ZM Payload
					/// 
					/// The message string will appear as: SS3 LT1 ""
					///
					/// This information contained here is in blocks of sub messages
					///   SS3  => Source connected [SS] followed by source number 1 to 8
					///   LT1  => Line Text        [LT] with 1 for Line 1, 2 for Line 2
					///   ""   => Actual text representation
					///   
					/// </summary>

					Logger.DebugFormat("{0}: Recived Zone {1} Message: Data Block is {2}", this.DeviceName, ZoneID, statusMessage);

					break;
			}

		}

		/// <summary>
		///  Send the command to the matrix
		/// </summary>
		/// <param name="data"></param>
		private void sendCommand(string data)
		{
			Logger.DebugFormat("{0}: Sending Command - {1}", this.DeviceName, data);

			_serial.Send(data + TX_TERMINATOR,300);
		}




		#endregion


		#region Main Driver Methods and Properties

		///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
		// 
		// The Following is a complete list of all the properties, methods and events exposed by this driver
		//
		// Events
		//   *All Zones Off Changed
		//   Zone Balance Level Changed
		//   Zone Bass Level Changed
		//   *Zone DND State Changed
		//   *Zone Loudness State Changed
		//   *Zone Name Changed
		//   *Zone Party Mode State Changed
		//   Zone Power State Changed
		//   Zone Source Changed
		//   *Zone Source Name Changed
		//   *Zone Shared State Changed
		//   *Zone Text Message Changed
		//   Zone Treble Level Changed
		//   Zone Volume Level Changed
		//   Zone Mute State Changed
		//   *Keypad Button Pressed
		//   *Display Feedback Received
		//
		// Instance Properties
		//   *AllZonesOff
		//   ZoneBalanceLevels
		//   ZoneBassLevels
		//   *ZoneDNDStates
		//   *ZoneLoudnessStates
		//   ZoneNames
		//   *ZonePartyModeStates
		//   ZonePowerStates
		//   ZoneSources
		//   *ZoneSourceNames
		//   SourceNames
		//   *ZoneSharedStates
		//   *ZoneTextMessages
		//   ZoneTrebleLevels
		//   ZoneVolumes
		//   ZoneMuteStates
		//
		// Instance Methods
		//   *GetAllZoneData
		//   *GetZoneData ( Number )
		//   TurnAllZonesOff
		//   TurnAllZonesOn
		//   ToggleZonePower ( Number )
		//   SetZonePowerState ( Number, Boolean )
		//   TurnZoneOn ( Number )
		//   TurnZoneOff ( Number )
		//   *DisplayAllMessage ( String, Number, Number )
		//   *DisplayZoneMessage ( Number, String, Number, Number )
		//   SetZoneVolume ( Number, Number )
		//   IncrementZoneVolume ( Number )
		//   DecrementZoneVolume ( Number )
		//   SetZoneSource ( Number, Number )
		//   CycleZoneSource ( Number )
		//   SetZoneBass ( Number, Number )
		//   IncrementZoneBass ( Number )
		//   DecrementZoneBass ( Number )
		//   SetZoneTreble ( Number, Number )
		//   IncrementZoneTreble ( Number )
		//   DecrementZoneTreble ( Number )
		//   SetZoneBalance ( Number, Number )
		//   IncrementZoneBalance ( Number )
		//   DecrementZoneBalance ( Number )
		//   *SetZoneTurnOnVolume ( Number, Number )
		//   *IncrementZoneTurnOnVolume ( Number )
		//   *DecrementZoneTurnOnVolume ( Number )
		//   *SetZoneDoNotDisturb ( Number, Boolean )
		//   *ToggleZoneDoNotDisturb ( Number )
		//   *SetZoneLoudness ( Number, Boolean )
		//   *ToggleZoneLoudness ( Number )
		//   *SetZonePartyMode ( Number, Number )
		//   MuteAllZones
		//   UnmuteAllZones
		//   ToggleZoneMute ( Number )
		//   SetZoneMuteState ( Number, Boolean )
		//   MuteZone ( Number )
		//   UnmuteZone ( Number )






		///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
		// Misc Device Functions
		//
		// The following method will allow us to directly access the control interface on the matrix
		//
		// We will create the following Functions:
		//
		//  Public Properties
		//      public IScriptArray ZoneNames
		//  Public Methods
		//      public void SendRawCommand(ScriptString rawCommand)
		//

		#region Misc Device Functions


		/// <summary>
		/// Expose a Property Array of the current zone names on the Matrix
		/// </summary>
		[ScriptObjectPropertyAttribute("Zone Names", "Gets the name of all zones.", "The {NAME} zone name for zone #{INDEX|1}.", null)]
		public IScriptArray ZoneNames
		{
			get
			{
				// In the 1st constructor parameter, _zoneNames is of type string[].
				// In the 2nd constructor parameter, 1 is the index of the 1st element in the create ScriptArray.
                return new ScriptArrayMarshalByValue(_zoneNames, 1);
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
		// Zone Power
		//
		// The following methods will allow us to manage the power on the physical matrix zones
		//
		// We will create the following Functions:
		//
		//  Private Methods
		//      private void setZoneVolume(int zoneNumber, int volume)
		//  
		//  Public Properties
		//      public IScriptArray ZonePowerStates
		// 
		//  Public Methods
		//      public void SetZonePower(ScriptNumber zoneNumber, ScriptBoolean power)
		//      public void TurnZoneOff(ScriptNumber zoneNumber)
		//      public void TurnZoneOn(ScriptNumber zoneNumber)
		//      public void ToggleZonePower(ScriptNumber zoneNumber)
		//      public void TurnAllZonesOff()
		//      public void TurnAllZonesOn()
		//

		#region Zone Power

		/// <summary>
		/// Private Function to actaully instruct the Matrix to change the Power on a Zone 
		/// </summary>
		/// <param name="zoneNumber">The Zone Number which we will be changing the Volume on</param>
		/// <param name="volume">The new Power setting we will be applying (On or Off)</param>
		private void setZonePower(int zoneNumber, bool power)
		{
			// First, ensure that a valid Zone number is provided based on the number of actually configured zones on the matrix

			if ((zoneNumber < 1) || (zoneNumber > _configuredZoneCount))
				throw new ArgumentOutOfRangeException("ZoneNumber", "The zone number must be between 1 and " + _configuredZoneCount);

			// Convert the Power setting from Boolean On/Off|True/False to a string representation
			string devicepower;
			if (power)
				devicepower = "1";
			else
				devicepower = "0";


			// Record to the Log that we are about to action a Power Change request
			Logger.InfoFormat("{0}: Zone {1} Set Power to {2}", this.DeviceName, zoneNumber, power);

			// Format the Devices Command String with the paramaters zoneNumber and Power, and send the command.
			//      !{z#}PR{v#}+  
			sendCommand("!" + zoneNumber + "PR" + devicepower + "+");
			sendCommand("!" + zoneNumber + "ZD+");
		}



		/// <summary>
		/// Expose a Property Array of current Zone Power Status. All Power Settings are exposed as True/False
		/// </summary>
		[ScriptObjectPropertyAttribute("Zone Power", "Gets or sets the current power setting for the zone. The setting is On/Off or True/False.", "the {NAME} power for zone #{INDEX|1}", "Set {NAME} zone #{INDEX|1} power  to #{VALUE|false|On|Off}", typeof(ScriptBoolean), 1, MaxZoneCount, "ZoneNames")]
		[SupportsDriverPropertyBinding("Zone Power State Changed", "Occurs when the current power setting for the zone changes.")]
		public IScriptArray ZonePowerStates
		{
			get
			{
				// If you have an bool[] array dedicated to power states then you could do this:
				return new ScriptArrayMarshalByReference(_zonePowerStates, new ScriptArraySetBooleanCallback(setZonePower), 1);
			}
		}




		/// <summary>
		/// Exposes a method to set the Power on a specified zone in the Matrix
		/// </summary>
		/// <param name="zoneNumber">The Zone Number which we will be Changing the Power on</param>
		/// <param name="power">The new Power Setting we will be applying (True/False)</param>
		[ScriptObjectMethodAttribute("Set Zone Power", "Change the Power Status for a Zone.", "Set the Power for {NAME} zone {PARAM|0|1} to {PARAM|1|false|on|off}.")]
		[ScriptObjectMethodParameterAttribute("zoneNumber", "The zone number.", 1, MaxZoneCount, "ZoneNames")]
		[ScriptObjectMethodParameter("Power", "Set the Power Status to ON or OFF", new string[] { "On", "Off" })]
		public void SetZonePower(ScriptNumber zoneNumber, ScriptBoolean power)
		{
			setZonePower((int)zoneNumber, (bool)power);
		}




		/// <summary>
		/// Exposes a method to Power Off a specified zone in the Matrix
		/// </summary>
		/// <param name="zoneNumber">The Zone Number we changing the power on</param>
		[ScriptObjectMethodAttribute("Turn Off Zone Power", "Turn Off the Power for a Zone.", "Turn Off the Power for {NAME} zone {PARAM|0|1}")]
		[ScriptObjectMethodParameterAttribute("zoneNumber", "The zone number.", 1, MaxZoneCount, "ZoneNames")]
		public void TurnZoneOff(ScriptNumber zoneNumber)
		{
			// First, ensure that a valid Zone number is provided based on the number of actually configured zones on the matrix
			if ((zoneNumber < 1) || (zoneNumber > _configuredZoneCount))
				throw new ArgumentOutOfRangeException("ZoneNumber", "The zone number must be between 1 and " + _configuredZoneCount);

			// Record to the Log that we are about to action a Power Change request
			Logger.InfoFormat("{0}: Zone {1} Power Off", this.DeviceName, (int)zoneNumber);

			// Format the Devices Command String with the paramater zoneNumber, and send the command.
			//      !{z#}PR0+  
			sendCommand("!" + (int)zoneNumber + "PR0+");
			sendCommand("!" + (int)zoneNumber + "ZD+");
		}




		/// <summary>
		/// Exposes a method to Power On a specified zone in the Matrix
		/// </summary>
		/// <param name="zoneNumber">The Zone Number we changing the power on</param>
		[ScriptObjectMethodAttribute("Turn On Zone Power", "Turn On the Power for a Zone.", "Turn On the Power for {NAME} zone {PARAM|0|1}")]
		[ScriptObjectMethodParameterAttribute("zoneNumber", "The zone number.", 1, MaxZoneCount, "ZoneNames")]
		public void TurnZoneOn(ScriptNumber zoneNumber)
		{
			// First, ensure that a valid Zone number is provided based on the number of actually configured zones on the matrix
			if ((zoneNumber < 1) || (zoneNumber > _configuredZoneCount))
				throw new ArgumentOutOfRangeException("ZoneNumber", "The zone number must be between 1 and " + _configuredZoneCount);

			// Record to the Log that we are about to action a Power Change request
			Logger.InfoFormat("{0}: Zone {1} Power On", this.DeviceName, (int)zoneNumber);

			// Format the Devices Command String with the paramater zoneNumber, and send the command.
			//      !{z#}PR1+  
			sendCommand("!" + (int)zoneNumber + "PR1+");
			sendCommand("!" + (int)zoneNumber + "ZD+");
		}




		/// <summary>
		/// Exposes a method to Toggle the Power On a specified zone in the Matrix
		/// </summary>
		/// <param name="zoneNumber">The Zone Number we changing the power on</param>
		[ScriptObjectMethodAttribute("Toggle Zone Power", "Toggle the Power for a Zone.", "Toggle the Power for {NAME} zone {PARAM|0|1}")]
		[ScriptObjectMethodParameterAttribute("zoneNumber", "The zone number.", 1, MaxZoneCount, "ZoneNames")]
		public void ToggleZonePower(ScriptNumber zoneNumber)
		{
			// First, ensure that a valid Zone number is provided based on the number of actually configured zones on the matrix
			if ((zoneNumber < 1) || (zoneNumber > _configuredZoneCount))
				throw new ArgumentOutOfRangeException("ZoneNumber", "The zone number must be between 1 and " + _configuredZoneCount);

			// Record to the Log that we are about to action a Power Change request
			Logger.InfoFormat("{0}: Zone {1} Power Toggle", this.DeviceName, (int)zoneNumber);

			// Format the Devices Command String with the paramater zoneNumber, and send the command.
			//      !{z#}PT+  
			sendCommand("!" + (int)zoneNumber + "PT+");
			sendCommand("!" + (int)zoneNumber + "ZD+");
		}




		/// <summary>
		/// Exposes a method to Power Off all the zones on the Matrix
		/// </summary>
		[ScriptObjectMethodAttribute("Power Off All Zones", "Power Down all Zone.", "Power Off all Zones")]
		public void TurnAllZonesOff()
		{
			// Loop trought each of the zones on the device, and call the Method to power it off

			for (int zone = 1; zone <= _configuredZoneCount; zone++)
			{
				TurnZoneOff(new ScriptNumber(zone));
			}
		}




		/// <summary>
		/// Exposes a method to Power On all the zones on the Matrix
		/// </summary>
		[ScriptObjectMethodAttribute("Power On All Zones", "Power Up all Zone.", "Power On all Zones")]
		public void TurnAllZonesOn()
		{
			// Loop trought each of the zones on the device, and call the Method to power it on

			for (int zone = 1; zone <= _configuredZoneCount; zone++)
			{
				TurnZoneOn(new ScriptNumber(zone));
			}
		}


		#endregion


		///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
		// Zone Mute
		//
		// The following methods will allow us to manage the Muting on the physical matrix zones
		//
		//  Private Methods
		//      private void setZoneMute(int zoneNumber, bool mute)
		//  
		//  Public Properties
		//      public IScriptArray ZoneMuteStates
		// 
		//  Public Methods
		//      public void SetZoneMute(ScriptNumber zoneNumber, ScriptBoolean mute)
		//      public void MuteZone(ScriptNumber zoneNumber)
		//      public void UnmuteZone(ScriptNumber zoneNumber)
		//      public void ToggleZoneMute(ScriptNumber zoneNumber)
		//      public void MuteAllZones()
		//      public void UnMuteAllZones()

		#region Zone Mute

		/// <summary>
		/// Private Function to actaully instruct the Matrix to change the Mute on a Zone 
		/// </summary>
		/// <param name="zoneNumber">The Zone Number which we will be changing the Mute on</param>
		/// <param name="volume">The new Mute setting we will be applying (On or Off)</param>
		private void setZoneMute(int zoneNumber, bool mute)
		{
			// First, ensure that a valid Zone number is provided based on the number of actually configured zones on the matrix

			if ((zoneNumber < 1) || (zoneNumber > _configuredZoneCount))
				throw new ArgumentOutOfRangeException("ZoneNumber", "The zone number must be between 1 and " + _configuredZoneCount);

			// Convert the Power setting from Boolean On/Off|True/False to a string representation
			string devicemute;
			if (mute)
				devicemute = "1";
			else
				devicemute = "0";


			// Record to the Log that we are about to action a Power Change request
			Logger.InfoFormat("{0}: Zone {1} Set Mute to {2}", this.DeviceName, zoneNumber, mute);

			// Format the Devices Command String with the paramaters zoneNumber and Mute, and send the command.
			//      !{z#}MU{v#}+  
			sendCommand("!" + zoneNumber + "MU" + devicemute + "+");
			sendCommand("!" + zoneNumber + "ZD+");
		}



		/// <summary>
		/// Expose a Property Array of current Zone Mute Status. All Mute Settings are exposed as True/False
		/// </summary>
		[ScriptObjectPropertyAttribute("Zone Mute States", "Gets or sets the current mute setting for the zone. The setting is On/Off or True/False.", "the {NAME} mute for zone #{INDEX|1}", "Set {NAME} zone #{INDEX|1} mute to #{VALUE|false|On|Off}", typeof(ScriptBoolean), 1, MaxZoneCount, "ZoneNames")]
		[SupportsDriverPropertyBinding("Zone Mute State Changed", "Occurs when the current mute setting for the zone changes.")]
		public IScriptArray ZoneMuteStates
		{
			get
			{
				// If you have an bool[] array dedicated to power states then you could do this:
				return new ScriptArrayMarshalByReference(_zonePowerStates, new ScriptArraySetBooleanCallback(setZoneMute), 1);
			}
		}




		/// <summary>
		/// Exposes a method to set the Mute on a specified zone in the Matrix
		/// </summary>
		/// <param name="zoneNumber">The Zone Number which we will be Changing the Mute on</param>
		/// <param name="mute">The new Mute Setting we will be applying (True/False)</param>
		[ScriptObjectMethodAttribute("Set Zone Mute", "Change the Mute Status for a Zone.", "Set the Mute for {NAME} zone {PARAM|0|1} to {PARAM|1|false|on|off}.")]
		[ScriptObjectMethodParameterAttribute("zoneNumber", "The zone number.", 1, MaxZoneCount, "ZoneNames")]
		[ScriptObjectMethodParameter("Mute", "Set the Mute Status to ON or OFF", new string[] { "On", "Off" })]
		public void SetZoneMute(ScriptNumber zoneNumber, ScriptBoolean mute)
		{
			setZoneMute((int)zoneNumber, (bool)mute);
		}




		/// <summary>
		/// Exposes a method to Mute a specific zone in the matrix
		/// </summary>
		/// <param name="zoneNumber">The Zone Number which we will be Muting</param>
		[ScriptObjectMethodAttribute("Mute Zone", "Turn On the Mute Status on a Zone.", "Enable Muteing on {NAME} zone {PARAM|0|1}")]
		[ScriptObjectMethodParameterAttribute("zoneNumber", "The zone number.", 1, MaxZoneCount, "ZoneNames")]
		public void MuteZone(ScriptNumber zoneNumber)
		{
			// First, ensure that a valid Zone number is provided based on the number of actually configured zones on the matrix
			if ((zoneNumber < 1) || (zoneNumber > _configuredZoneCount))
				throw new ArgumentOutOfRangeException("ZoneNumber", "The zone number must be between 1 and " + _configuredZoneCount);

			// Record to the Log that we are about to action a Power Change request
			Logger.InfoFormat("{0}: Zone {1} Mute On", this.DeviceName, (int)zoneNumber);

			// Format the Devices Command String with the paramater zoneNumber, and send the command.
			//      !{z#}MU1+  
			sendCommand("!" + (int)zoneNumber + "MU1+");
			sendCommand("!" + (int)zoneNumber + "ZD+");
		}




		/// <summary>
		/// Exposes a method to Unmute a specific zone in the matrix
		/// </summary>
		/// <param name="zoneNumber">The Zone Number which we will be Unmuting</param>
		[ScriptObjectMethodAttribute("Unmute Zone", "Turn Off the Mute Status on a Zone.", "Disable Muteing on {NAME} zone {PARAM|0|1}")]
		[ScriptObjectMethodParameterAttribute("zoneNumber", "The zone number.", 1, MaxZoneCount, "ZoneNames")]
		public void UnmuteZone(ScriptNumber zoneNumber)
		{
			// First, ensure that a valid Zone number is provided based on the number of actually configured zones on the matrix
			if ((zoneNumber < 1) || (zoneNumber > _configuredZoneCount))
				throw new ArgumentOutOfRangeException("ZoneNumber", "The zone number must be between 1 and " + _configuredZoneCount);

			// Record to the Log that we are about to action a Power Change request
			Logger.InfoFormat("{0}: Zone {1} Mute Off", this.DeviceName, (int)zoneNumber);

			// Format the Devices Command String with the paramater zoneNumber, and send the command.
			//      !{z#}MU0+  
			sendCommand("!" + (int)zoneNumber + "MU0+");
			sendCommand("!" + (int)zoneNumber + "ZD+");
		}




		/// <summary>
		/// Exposes a method to Toggle the Mute setting on a specific zone in the matrix
		/// </summary>
		/// <param name="zoneNumber">The Zone Number which we will be Changing</param>
		[ScriptObjectMethodAttribute("Toggle Zone Mute", "Toggle the Mute Status for a Zone.", "Toggle Mute for {NAME} zone {PARAM|0|1}")]
		[ScriptObjectMethodParameterAttribute("zoneNumber", "The zone number.", 1, MaxZoneCount, "ZoneNames")]
		public void ToggleZoneMute(ScriptNumber zoneNumber)
		{
			// First, ensure that a valid Zone number is provided based on the number of actually configured zones on the matrix
			if ((zoneNumber < 1) || (zoneNumber > _configuredZoneCount))
				throw new ArgumentOutOfRangeException("ZoneNumber", "The zone number must be between 1 and " + _configuredZoneCount);

			// Record to the Log that we are about to action a Power Change request
			Logger.InfoFormat("{0}: Zone {1} Mute Toggle", this.DeviceName, (int)zoneNumber);

			// Format the Devices Command String with the paramater zoneNumber, and send the command.
			//      !{z#}MT+  
			sendCommand("!" + (int)zoneNumber + "MT+");
			sendCommand("!" + (int)zoneNumber + "ZD+");
		}




		/// <summary>
		/// Exposes a method to Mute all zones on the matrix
		/// </summary>
		[ScriptObjectMethodAttribute("Mute All Zones", "Enable Mute on all Zones.", "Mute all Zones")]
		public void MuteAllZones()
		{
			// Loop trought each of the zones on the device, and call the Method to Mute it

			for (int zone = 1; zone <= _configuredZoneCount; zone++)
			{
				MuteZone(new ScriptNumber(zone));
			}
		}



		/// <summary>
		/// Exposes a method to Unmute all zones on the matrix
		/// </summary>
		[ScriptObjectMethodAttribute("Unmute All Zones", "Disable Mute on all Zones.", "Unmute all Zones")]
		public void UnMuteAllZones()
		{
			// Loop trought each of the zones on the device, and call the Method to unmute it

			for (int zone = 1; zone <= _configuredZoneCount; zone++)
			{
				UnmuteZone(new ScriptNumber(zone));
			}
		}


		#endregion


		///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
		// Zone Volume
		//
		// The following methods will allow us to manage the volume on the physical matrix
		//
		// We will create the following Functions:
		//
		//  Private Methods
		//      private void setZoneVolume(int zoneNumber, int volume)
		//  
		//  Public Properties
		//      public IScriptArray ZoneVolumes
		// 
		//  Public Methods
		//      public void SetZoneVolume(ScriptNumber zoneNumber, ScriptNumber volume)
		//      public void IncrementZoneVolume(ScriptNumber zoneNumber)
		//      public void DecrementZoneVolume(ScriptNumber zoneNumber)
		//

		#region Zone Volume

		/// <summary>
		/// Private Function to actaully instruct the Matrix to change the Volume on a Zone 
		/// </summary>
		/// <param name="zoneNumber">The Zone Number which we will be changing the Volume on</param>
		/// <param name="volume">The new Volume Level we will be applying (Range is 0 to 100)</param>
		private void setZoneVolume(int zoneNumber, int volume)
		{
			// First, ensure that a valid Zone number is provided based on the number of actually configured zones on the matrix

			if ((zoneNumber < 1) || (zoneNumber > _configuredZoneCount))
				throw new ArgumentOutOfRangeException("ZoneNumber", "The zone number must be between 1 and " + _configuredZoneCount);

			// Next, ensure that the volume level requested is within the range required of 0 to 100, round back if required.
			if (volume < 0) volume = 0;
			if (volume > 100) volume = 100;

			// Range on the Matrix for Volume is 0 to 38, which we will be exposing to the users in the scale of 0 to 100
			// to achieve this, we will do a simple match conversion when writing and reading the volume level from the device
			// as wer store and retrive the base 100 result. Note that we will alway use Absoute numbers to communictate with the device
			//
			// Example - 
			//  If the user sets volume to 100, we calculate the true volume to be 38 
			//  If the user sets volume to 50, we calculate the true volume to be 19
			//
			// Forumla -  
			//  devicevol = volume * 0.38

			byte devicevol = (byte)Math.Abs(volume * 0.38);

			// Record to the Log that we are about to action a Volume Change request
			Logger.InfoFormat("{0}: Zone {1} Set Volume to {2}", this.DeviceName, zoneNumber, volume);

			// Format the Devices Command String with the paramaters zoneNumber and Volume, and send the command.
			//      !{z#}VO{v#}+  
			sendCommand("!" + zoneNumber + "VO" + devicevol + "+");
			sendCommand("!" + zoneNumber + "ZD+");
		}



		/// <summary>
		/// Expose a Property Array of current Zone Volume Levels. All Volume levels are exposed in the Range 0 to 100
		/// </summary>
		[ScriptObjectPropertyAttribute("Zone Volume Levels", "Gets or sets the current volume setting for the player zones. The scale is 0 to 100.", 0, 100, "the {NAME} volume for zone #{INDEX|1}", "Set {NAME} zone #{INDEX|1} volume to #{VALUE|100}", typeof(ScriptNumber), 1, MaxZoneCount, "ZoneNames")]
		[SupportsDriverPropertyBinding("Zone Volume Level Changed", "Occurs when the current volume setting for the player changes.")]
		public IScriptArray ZoneVolumes
		{
			get
			{
				// If you have an int[] array dedicated to volumes then you could do this:
				return new ScriptArrayMarshalByReference(_zoneVolumeLevels, new ScriptArraySetInt32Callback(setZoneVolume), 1);
			}
		}



		/// <summary>
		/// Exposes a method to set the volume on a specified zone in the Matrix
		/// </summary>
		/// <param name="zoneNumber">The Zone Number which we will be Changing the Volume on</param>
		/// <param name="volume">The new Volume Level we will be applying (Range is 0 to 100)</param>
		[ScriptObjectMethodAttribute("Set Zone Volume", "Change the Volume for a Zone.", "Set the volume for {NAME} zone #{PARAM|0|1} to {PARAM|1|10}.")]
		[ScriptObjectMethodParameterAttribute("zoneNumber", "The zone number.", 1, MaxZoneCount, "ZoneNames")]
		[ScriptObjectMethodParameterAttribute("Volume", "Volume (0=Quietest, 100=Loudest)", 0, 100)]
		public void SetZoneVolume(ScriptNumber zoneNumber, ScriptNumber volume)
		{
			setZoneVolume((int)zoneNumber, (int)volume);
		}



		/// <summary>
		/// Exposes a method to Increment the volume on a specified zone in the Matrix
		/// </summary>
		/// <param name="zoneNumber">The Zone Number we will be Incrementing the Volume on</param>
		[ScriptObjectMethodAttribute("Increment Zone Volume", "Increment the Volume for a Zone.", "Increment the Voluem for {NAME} zone {PARAM|0|1}")]
		[ScriptObjectMethodParameterAttribute("zoneNumber", "The zone number.", 1, MaxZoneCount, "ZoneNames")]
		public void IncrementZoneVolume(ScriptNumber zoneNumber)
		{
			// First, ensure that a valid Zone number is provided based on the number of actually configured zones on the matrix
			if ((zoneNumber < 1) || (zoneNumber > _configuredZoneCount))
				throw new ArgumentOutOfRangeException("ZoneNumber", "The zone number must be between 1 and " + _configuredZoneCount);

			// We will use the Matrix Command for Incrementing the Zones Volume, therfore we do not need to be concerned with calculating
			// the correct base range, or ensuring that the volume level is not out of range as the Matrix will manage this directly.

			// Record to the Log that we are about to action a Volume Change request
			Logger.InfoFormat("{0}: Zone {1} Increment Volume", this.DeviceName, (int)zoneNumber);

			// Format the Devices Command String with the paramater zoneNumber, and send the command.
			//      !{z#}VI+  
			sendCommand("!" + (int)zoneNumber + "VI+");
			sendCommand("!" + (int)zoneNumber + "ZD+");
		}



		/// <summary>
		/// Exposes a method to Decrement the volume on a specified zone in the Matrix
		/// </summary>
		/// <param name="zoneNumber">The Zone Number we will be Decrementing the Volume on</param>
		[ScriptObjectMethodAttribute("Decrement Zone Volume", "Decrement the Volume for a Zone.", "Decrement the Voluem for {NAME} zone {PARAM|0|1}")]
		[ScriptObjectMethodParameterAttribute("zoneNumber", "The zone number.", 1, MaxZoneCount, "ZoneNames")]
		public void DecrementZoneVolume(ScriptNumber zoneNumber)
		{
			// First, ensure that a valid Zone number is provided based on the number of actually configured zones on the matrix
			if ((zoneNumber < 1) || (zoneNumber > _configuredZoneCount))
				throw new ArgumentOutOfRangeException("ZoneNumber", "The zone number must be between 1 and " + _configuredZoneCount);

			// We will use the Matrix Command for Decrementing the Zones Volume, therfore we do not need to be concerned with calculating
			// the correct base range, or ensuring that the volume level is not out of range as the Matrix will manage this directly.

			// Record to the Log that we are about to action a Volume Change request
			Logger.InfoFormat("{0}: Zone {1} Decrement Volume", this.DeviceName, (int)zoneNumber);

			// Format the Devices Command String with the paramater zoneNumber, and send the command.
			//      !{z#}VD+  
			sendCommand("!" + (int)zoneNumber + "VD+");
			sendCommand("!" + (int)zoneNumber + "ZD+");
		}


		#endregion




		///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
		// Zone Bass
		//
		// The following methods will allow us to manage the volume on the physical matrix
		//
		// We will create the following Functions:
		//
		//  Private Methods
		//      private void setZoneBass(int zoneNumber, int bass)
		//  
		//  Public Properties
		//      public IScriptArray ZoneBassLevels
		// 
		//  Public Methods
		//      public void SetZoneBass(ScriptNumber zoneNumber, ScriptNumber bass)
		//      public void IncrementZoneBass(ScriptNumber zoneNumber)
		//      public void DecrementZoneBass(ScriptNumber zoneNumber)
		//

		#region Zone Bass

		/// <summary>
		/// Private Function to actaully instruct the Matrix to change the Volume on a Bass 
		/// </summary>
		/// <param name="zoneNumber">The Zone Number which we will be changing the Bass on</param>
		/// <param name="volume">The new Volume Level we will be applying (Range is 0 to 100)</param>
		private void setZoneBass(int zoneNumber, int bass)
		{
			// First, ensure that a valid Zone number is provided based on the number of actually configured zones on the matrix

			if ((zoneNumber < 1) || (zoneNumber > _configuredZoneCount))
				throw new ArgumentOutOfRangeException("ZoneNumber", "The zone number must be between 1 and " + _configuredZoneCount);

			// Next, ensure that the volume level requested is within the range required of 0 to 100, round back if required.
			if (bass < 0) bass = 0;
			if (bass > 100) bass = 100;

			// Range on the Matrix for Volume is 0 to 14, which we will be exposing to the users in the scale of 0 to 100
			// to achieve this, we will do a simple match conversion when writing and reading the volume level from the device
			// as wer store and retrive the base 100 result. Note that we will alway use Absoute numbers to communictate with the device
			//
			// Example - 
			//  If the user sets volume to 100, we calculate the true volume to be 14 
			//  If the user sets volume to 50, we calculate the true volume to be 7
			//
			// Forumla -  
			//  devicevol = volume * 0.14

			byte devicebass = (byte)Math.Abs(bass * 0.14);

			// Record to the Log that we are about to action a Volume Change request
			Logger.InfoFormat("{0}: Zone {1} Set Bass to {2}", this.DeviceName, zoneNumber, bass);

			// Format the Devices Command String with the paramaters zoneNumber and Volume, and send the command.
			//      !{z#}VO{v#}+  
			sendCommand("!" + zoneNumber + "BS" + devicebass + "+");
			sendCommand("!" + zoneNumber + "ZD+");
		}



		/// <summary>
		/// Expose a Property Array of current Zone Bass Levels. All Volume levels are exposed in the Range 0 to 100
		/// </summary>
		[ScriptObjectPropertyAttribute("Zone Bass Levels", "Gets or sets the current bass setting for the player zones. The scale is 0 to 100.", 0, 100, "the {NAME} bass for zone #{INDEX|1}", "Set {NAME} zone #{INDEX|1} bass to {VALUE|100}", typeof(ScriptNumber), 1, MaxZoneCount, "ZoneNames")]
		[SupportsDriverPropertyBinding("Zone Bass Level Changed", "Occurs when the current bass setting for the matrix changes.")]
		public IScriptArray ZoneBassLevels
		{
			get
			{
				// If you have an int[] array dedicated to volumes then you could do this:
				return new ScriptArrayMarshalByReference(_zoneVolumeLevels, new ScriptArraySetInt32Callback(setZoneBass), 1);
			}
		}



		/// <summary>
		/// Exposes a method to set the Bass on a specified zone in the Matrix
		/// </summary>
		/// <param name="zoneNumber">The Zone Number which we will be Changing the Bass on</param>
		/// <param name="volume">The new Volume Level we will be applying (Range is 0 to 100)</param>
		[ScriptObjectMethodAttribute("Set Zone Bass", "Change the Bass for a Zone.", "Set the bass for {NAME} zone #{PARAM|0|1} to {PARAM|1|10}.")]
		[ScriptObjectMethodParameterAttribute("zoneNumber", "The zone number.", 1, MaxZoneCount, "ZoneNames")]
		[ScriptObjectMethodParameterAttribute("Bass", "Volume (0=Quietest, 100=Loudest)", 0, 100)]
		public void SetZoneBass(ScriptNumber zoneNumber, ScriptNumber bass)
		{
			setZoneVolume((int)zoneNumber, (int)bass);
		}



		/// <summary>
		/// Exposes a method to Increment the Bass on a specified zone in the Matrix
		/// </summary>
		/// <param name="zoneNumber">The Zone Number we will be Incrementing the Bass on</param>
		[ScriptObjectMethodAttribute("Increment Zone Bass", "Increment the Bass for a Zone.", "Increment the Bass for {NAME} zone {PARAM|0|1}")]
		[ScriptObjectMethodParameterAttribute("zoneNumber", "The zone number.", 1, MaxZoneCount, "ZoneNames")]
		public void IncrementZoneBass(ScriptNumber zoneNumber)
		{
			// First, ensure that a valid Zone number is provided based on the number of actually configured zones on the matrix
			if ((zoneNumber < 1) || (zoneNumber > _configuredZoneCount))
				throw new ArgumentOutOfRangeException("ZoneNumber", "The zone number must be between 1 and " + _configuredZoneCount);

			// We will use the Matrix Command for Incrementing the Zones Volume, therfore we do not need to be concerned with calculating
			// the correct base range, or ensuring that the volume level is not out of range as the Matrix will manage this directly.

			// Record to the Log that we are about to action a Bass Change request
			Logger.InfoFormat("{0}: Zone {1} Increment Bass", this.DeviceName, (int)zoneNumber);

			// Format the Devices Command String with the paramater zoneNumber, and send the command.
			//      !{z#}VI+  
			sendCommand("!" + (int)zoneNumber + "BI+");
			sendCommand("!" + (int)zoneNumber + "ZD+");
		}



		/// <summary>
		/// Exposes a method to Decrement the Bass on a specified zone in the Matrix
		/// </summary>
		/// <param name="zoneNumber">The Zone Number we will be Decrementing the Bass on</param>
		[ScriptObjectMethodAttribute("Decrement Zone Bass", "Decrement the Bass for a Zone.", "Decrement the Bass for {NAME} zone {PARAM|0|1}")]
		[ScriptObjectMethodParameterAttribute("zoneNumber", "The zone number.", 1, MaxZoneCount, "ZoneNames")]
		public void DecrementZoneBass(ScriptNumber zoneNumber)
		{
			// First, ensure that a valid Zone number is provided based on the number of actually configured zones on the matrix
			if ((zoneNumber < 1) || (zoneNumber > _configuredZoneCount))
				throw new ArgumentOutOfRangeException("ZoneNumber", "The zone number must be between 1 and " + _configuredZoneCount);

			// We will use the Matrix Command for Decrementing the Zones Volume, therfore we do not need to be concerned with calculating
			// the correct base range, or ensuring that the volume level is not out of range as the Matrix will manage this directly.

			// Record to the Log that we are about to action a Bass Change request
			Logger.InfoFormat("{0}: Zone {1} Decrement Bass", this.DeviceName, (int)zoneNumber);

			// Format the Devices Command String with the paramater zoneNumber, and send the command.
			//      !{z#}VD+  
			sendCommand("!" + (int)zoneNumber + "BD+");
			sendCommand("!" + (int)zoneNumber + "ZD+");
		}


		#endregion



		///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
		// Zone Treble
		//
		// The following methods will allow us to manage the volume on the physical matrix
		//
		// We will create the following Functions:
		//
		//  Private Methods
		//      private void setZoneTreble(int zoneNumber, int treble)
		//  
		//  Public Properties
		//      public IScriptArray ZoneTrebleLevels
		// 
		//  Public Methods
		//      public void SetZoneTreble(ScriptNumber zoneNumber, ScriptNumber treble)
		//      public void IncrementZoneTreble(ScriptNumber zoneNumber)
		//      public void DecrementZoneTreble(ScriptNumber zoneNumber)
		//

		#region Zone Treble

		/// <summary>
		/// Private Function to actaully instruct the Matrix to change the Volume on a Treble
		/// </summary>
		/// <param name="zoneNumber">The Zone Number which we will be changing the Treble on</param>
		/// <param name="volume">The new Treble Level we will be applying (Range is 0 to 100)</param>
		private void setZoneTreble(int zoneNumber, int treble)
		{
			// First, ensure that a valid Zone number is provided based on the number of actually configured zones on the matrix

			if ((zoneNumber < 1) || (zoneNumber > _configuredZoneCount))
				throw new ArgumentOutOfRangeException("ZoneNumber", "The zone number must be between 1 and " + _configuredZoneCount);

			// Next, ensure that the volume level requested is within the range required of 0 to 100, round back if required.
			if (treble < 0) treble = 0;
			if (treble > 100) treble = 100;

			// Range on the Matrix for Volume is 0 to 14, which we will be exposing to the users in the scale of 0 to 100
			// to achieve this, we will do a simple match conversion when writing and reading the volume level from the device
			// as wer store and retrive the base 100 result. Note that we will alway use Absoute numbers to communictate with the device
			//
			// Example - 
			//  If the user sets volume to 100, we calculate the true volume to be 14 
			//  If the user sets volume to 50, we calculate the true volume to be 7
			//
			// Forumla -  
			//  devicevol = volume * 0.14

			byte devicetreble = (byte)Math.Abs(treble * 0.14);

			// Record to the Log that we are about to action a Treble Change request
			Logger.InfoFormat("{0}: Zone {1} Set Treble to {2}", this.DeviceName, zoneNumber, treble);

			// Format the Devices Command String with the paramaters zoneNumber and Volume, and send the command.
			//      !{z#}VO{v#}+  
			sendCommand("!" + zoneNumber + "TR" + devicetreble + "+");
			sendCommand("!" + zoneNumber + "ZD+");
		}



		/// <summary>
		/// Expose a Property Array of current Zone Treble Levels. All Volume levels are exposed in the Range 0 to 100
		/// </summary>
		[ScriptObjectPropertyAttribute("Zone Treble Levels", "Gets or sets the current treble setting for the player zones. The scale is 0 to 100.", 0, 100, "the {NAME} treble for zone #{INDEX|1}", "Set {NAME} zone #{INDEX|1} treble to {VALUE|100}", typeof(ScriptNumber), 1, MaxZoneCount, "ZoneNames")]
		[SupportsDriverPropertyBinding("Zone Treble Level Changed", "Occurs when the current treble setting for the matrix changes.")]
		public IScriptArray ZoneTrebleLevels
		{
			get
			{
				// If you have an int[] array dedicated to volumes then you could do this:
				return new ScriptArrayMarshalByReference(_zoneVolumeLevels, new ScriptArraySetInt32Callback(setZoneTreble), 1);
			}
		}



		/// <summary>
		/// Exposes a method to set the Treble on a specified zone in the Matrix
		/// </summary>
		/// <param name="zoneNumber">The Zone Number which we will be Changing the Treble on</param>
		/// <param name="volume">The new Treble Level we will be applying (Range is 0 to 100)</param>
		[ScriptObjectMethodAttribute("Set Zone Treble", "Change the Treble for a Zone.", "Set the treble for {NAME} zone #{PARAM|0|1} to {PARAM|1|10}.")]
		[ScriptObjectMethodParameterAttribute("zoneNumber", "The zone number.", 1, MaxZoneCount, "ZoneNames")]
		[ScriptObjectMethodParameterAttribute("Treble", "Treble (0=Quietest, 100=Loudest)", 0, 100)]
		public void SetZoneTreble(ScriptNumber zoneNumber, ScriptNumber treble)
		{
			setZoneVolume((int)zoneNumber, (int)treble);
		}



		/// <summary>
		/// Exposes a method to Increment the Bass on a specified zone in the Matrix
		/// </summary>
		/// <param name="zoneNumber">The Zone Number we will be Incrementing the Bass on</param>
		[ScriptObjectMethodAttribute("Increment Zone Treble", "Increment the Treble for a Zone.", "Increment the treble for {NAME} zone {PARAM|0|1}")]
		[ScriptObjectMethodParameterAttribute("zoneNumber", "The zone number.", 1, MaxZoneCount, "ZoneNames")]
		public void IncrementZoneTreble(ScriptNumber zoneNumber)
		{
			// First, ensure that a valid Zone number is provided based on the number of actually configured zones on the matrix
			if ((zoneNumber < 1) || (zoneNumber > _configuredZoneCount))
				throw new ArgumentOutOfRangeException("ZoneNumber", "The zone number must be between 1 and " + _configuredZoneCount);

			// We will use the Matrix Command for Incrementing the Zones Volume, therfore we do not need to be concerned with calculating
			// the correct base range, or ensuring that the volume level is not out of range as the Matrix will manage this directly.

			// Record to the Log that we are about to action a Treble Change request
			Logger.InfoFormat("{0}: Zone {1} Increment Treble", this.DeviceName, (int)zoneNumber);

			// Format the Devices Command String with the paramater zoneNumber, and send the command.
			//      !{z#}VI+  
			sendCommand("!" + (int)zoneNumber + "TI+");
			sendCommand("!" + (int)zoneNumber + "ZD+");
		}



		/// <summary>
		/// Exposes a method to Decrement the Bass on a specified zone in the Matrix
		/// </summary>
		/// <param name="zoneNumber">The Zone Number we will be Decrementing the Bass on</param>
		[ScriptObjectMethodAttribute("Decrement Zone Treble", "Decrement the Treble for a Zone.", "Decrement the Treble for {NAME} zone {PARAM|0|1}")]
		[ScriptObjectMethodParameterAttribute("zoneNumber", "The zone number.", 1, MaxZoneCount, "ZoneNames")]
		public void DecrementZoneTreble(ScriptNumber zoneNumber)
		{
			// First, ensure that a valid Zone number is provided based on the number of actually configured zones on the matrix
			if ((zoneNumber < 1) || (zoneNumber > _configuredZoneCount))
				throw new ArgumentOutOfRangeException("ZoneNumber", "The zone number must be between 1 and " + _configuredZoneCount);

			// We will use the Matrix Command for Decrementing the Zones Volume, therfore we do not need to be concerned with calculating
			// the correct base range, or ensuring that the volume level is not out of range as the Matrix will manage this directly.

			// Record to the Log that we are about to action a Bass Change request
			Logger.InfoFormat("{0}: Zone {1} Decrement Bass", this.DeviceName, (int)zoneNumber);

			// Format the Devices Command String with the paramater zoneNumber, and send the command.
			//      !{z#}VD+  
			sendCommand("!" + (int)zoneNumber + "TD+");
			sendCommand("!" + (int)zoneNumber + "ZD+");
		}


		#endregion




		///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
		// Zone Balance
		//
		// The following methods will allow us to manage the volume on the physical matrix
		//
		// We will create the following Functions:
		//
		//  Private Methods
		//      private void setZoneBass(int zoneNumber, int bass)
		//  
		//  Public Properties
		//      public IScriptArray ZoneBassLevels
		// 
		//  Public Methods
		//      public void SetZoneBass(ScriptNumber zoneNumber, ScriptNumber bass)
		//      public void IncrementZoneBass(ScriptNumber zoneNumber)
		//      public void DecrementZoneBass(ScriptNumber zoneNumber)
		//

		#region Zone Balance

		/// <summary>
		/// Private Function to actaully instruct the Matrix to change the Volume on a Bass 
		/// </summary>
		/// <param name="zoneNumber">The Zone Number which we will be changing the Bass on</param>
		/// <param name="volume">The new Volume Level we will be applying (Range is 0 to 100)</param>
		private void setZoneBalance(int zoneNumber, int balance)
		{
			// First, ensure that a valid Zone number is provided based on the number of actually configured zones on the matrix

			if ((zoneNumber < 1) || (zoneNumber > _configuredZoneCount))
				throw new ArgumentOutOfRangeException("ZoneNumber", "The zone number must be between 1 and " + _configuredZoneCount);

			// Next, ensure that the volume level requested is within the range required of 0 to 100, round back if required.
			if (balance < 0) balance = 0;
			if (balance > 100) balance = 100;

			// Range on the Matrix for Volume is 0 to 14, which we will be exposing to the users in the scale of 0 to 100
			// to achieve this, we will do a simple match conversion when writing and reading the volume level from the device
			// as wer store and retrive the base 100 result. Note that we will alway use Absoute numbers to communictate with the device
			//
			// Example - 
			//  If the user sets volume to 100, we calculate the true volume to be 14 
			//  If the user sets volume to 50, we calculate the true volume to be 7
			//
			// Forumla -  
			//  devicevol = volume * 0.64

			byte devicebalance = (byte)Math.Abs(balance * 0.64);

			// Record to the Log that we are about to action a Volume Change request
			Logger.InfoFormat("{0}: Zone {1} Set Balance to {2}", this.DeviceName, zoneNumber, balance);

			// Format the Devices Command String with the paramaters zoneNumber and Volume, and send the command.
			//      !{z#}VO{v#}+  
			sendCommand("!" + zoneNumber + "BA" + devicebalance + "+");
			sendCommand("!" + zoneNumber + "ZD+");
		}



		/// <summary>
		/// Expose a Property Array of current Zone Bass Levels. All Volume levels are exposed in the Range 0 to 100
		/// </summary>
		[ScriptObjectPropertyAttribute("Zone Balance Levels", "Gets or sets the current balance setting for the player zones. The scale is 0 to 100.", 0, 100, "the {NAME} balance for zone #{INDEX|1}", "Set {NAME} zone #{INDEX|1} balance to {VALUE|100}", typeof(ScriptNumber), 1, MaxZoneCount, "ZoneNames")]
		[SupportsDriverPropertyBinding("Zone Balance Level Changed", "Occurs when the current balance setting for the matrix changes.")]
		public IScriptArray ZoneBalanceLevels
		{
			get
			{
				// If you have an int[] array dedicated to volumes then you could do this:
				return new ScriptArrayMarshalByReference(_zoneVolumeLevels, new ScriptArraySetInt32Callback(setZoneBalance), 1);
			}
		}



		/// <summary>
		/// Exposes a method to set the Bass on a specified zone in the Matrix
		/// </summary>
		/// <param name="zoneNumber">The Zone Number which we will be Changing the Bass on</param>
		/// <param name="volume">The new Volume Level we will be applying (Range is 0 to 100)</param>
		[ScriptObjectMethodAttribute("Set Zone Balance", "Change the Balance for a Zone.", "Set the balance for {NAME} zone #{PARAM|0|1} to {PARAM|1|10}.")]
		[ScriptObjectMethodParameterAttribute("zoneNumber", "The zone number.", 1, MaxZoneCount, "ZoneNames")]
		[ScriptObjectMethodParameterAttribute("Balance", "Balance (0=Quietest, 100=Loudest)", 0, 100)]
		public void SetZoneBalance(ScriptNumber zoneNumber, ScriptNumber balance)
		{
			setZoneVolume((int)zoneNumber, (int)balance);
		}



		/// <summary>
		/// Exposes a method to Increment the Bass on a specified zone in the Matrix
		/// </summary>
		/// <param name="zoneNumber">The Zone Number we will be Incrementing the Bass on</param>
		[ScriptObjectMethodAttribute("Increment Zone Balance", "Increment the Balance for a Zone.", "Increment the Balance for {NAME} zone {PARAM|0|1}")]
		[ScriptObjectMethodParameterAttribute("zoneNumber", "The zone number.", 1, MaxZoneCount, "ZoneNames")]
		public void IncrementZoneBalance(ScriptNumber zoneNumber)
		{
			// First, ensure that a valid Zone number is provided based on the number of actually configured zones on the matrix
			if ((zoneNumber < 1) || (zoneNumber > _configuredZoneCount))
				throw new ArgumentOutOfRangeException("ZoneNumber", "The zone number must be between 1 and " + _configuredZoneCount);

			// We will use the Matrix Command for Incrementing the Zones Volume, therfore we do not need to be concerned with calculating
			// the correct base range, or ensuring that the volume level is not out of range as the Matrix will manage this directly.

			// Record to the Log that we are about to action a Bass Change request
			Logger.InfoFormat("{0}: Zone {1} Increment Balance", this.DeviceName, (int)zoneNumber);

			// Format the Devices Command String with the paramater zoneNumber, and send the command.
			//      !{z#}VI+  
			sendCommand("!" + (int)zoneNumber + "BL+");
			sendCommand("!" + (int)zoneNumber + "ZD+");
		}



		/// <summary>
		/// Exposes a method to Decrement the Bass on a specified zone in the Matrix
		/// </summary>
		/// <param name="zoneNumber">The Zone Number we will be Decrementing the Bass on</param>
		[ScriptObjectMethodAttribute("Decrement Zone Balance", "Decrement the Balance for a Zone.", "Decrement the Balance for {NAME} zone {PARAM|0|1}")]
		[ScriptObjectMethodParameterAttribute("zoneNumber", "The zone number.", 1, MaxZoneCount, "ZoneNames")]
		public void DecrementZoneBalance(ScriptNumber zoneNumber)
		{
			// First, ensure that a valid Zone number is provided based on the number of actually configured zones on the matrix
			if ((zoneNumber < 1) || (zoneNumber > _configuredZoneCount))
				throw new ArgumentOutOfRangeException("ZoneNumber", "The zone number must be between 1 and " + _configuredZoneCount);

			// We will use the Matrix Command for Decrementing the Zones Volume, therfore we do not need to be concerned with calculating
			// the correct base range, or ensuring that the volume level is not out of range as the Matrix will manage this directly.

			// Record to the Log that we are about to action a Bass Change request
			Logger.InfoFormat("{0}: Zone {1} Decrement Balance", this.DeviceName, (int)zoneNumber);

			// Format the Devices Command String with the paramater zoneNumber, and send the command.
			//      !{z#}VD+  
			sendCommand("!" + (int)zoneNumber + "BR+");
			sendCommand("!" + (int)zoneNumber + "ZD+");
		}


		#endregion


		///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
		// Zone Sources
		//
		// The following methods will allow us to manage the volume on the physical matrix
		//
		// We will create the following Functions:
		//
		//  Private Methods
		//      private void setZoneSource(int zoneNumber, int source)
		//  
		//  Public Properties
		//      public IScriptArray ZoneSources
		//      public IScriptArray SourceNames
		// 
		//  Public Methods
		//      public void SetZoneVolume(ScriptNumber zoneNumber, ScriptNumber source)
		//      public void CycleZoneSource(ScriptNumber zoneNumber)
		//      public void DecrementZoneVolume(ScriptNumber zoneNumber)
		//

		#region Zone Sources


		/// <summary>
		/// Private Function to actaully instruct the Matrix to change the Source on a Zone 
		/// </summary>
		/// <param name="zoneNumber">The Zone Number which we will be changing the Source on</param>
		/// <param name="volume">The new Source Input we will be applying (Range is 1 to 8)</param>
		private void setZoneSource(int zoneNumber, int source)
		{
			// First, ensure that a valid Zone number is provided based on the number of actually configured zones on the matrix

			if ((zoneNumber < 1) || (zoneNumber > _configuredZoneCount))
				throw new ArgumentOutOfRangeException("ZoneNumber", "The zone number must be between 1 and " + _configuredZoneCount);

			// Next, ensure that the source number requested is within the range required of 1 to 8, round back if required.
			if (source < 1) source = 1;
			if (source > 8) source = 8;

			// Record to the Log that we are about to action a Source Change request
			Logger.InfoFormat("{0}: Zone {1} Set Source to {2}", this.DeviceName, zoneNumber, source);

			// Format the Devices Command String with the paramaters zoneNumber and Source, and send the command.
			//      !{z#}SS{v#}+  
			sendCommand("!" + zoneNumber + "SS" + source + "+");
			sendCommand("!" + (int)zoneNumber + "ZD+");
		}



		/// <summary>
		/// Expose a Property Array of current Zone Sourcess. All Sources are exposed in the Range 1 to 8
		/// </summary>
		[ScriptObjectPropertyAttribute("Zone Sources", "Gets or sets the current source setting for the zones.", "the {NAME} source for zone #{INDEX|1}", "Set {NAME} zone #{INDEX|1} source to {value|0}.", typeof(ScriptNumber), 1, MaxSourceCount, "ZoneNames")]
		[SupportsDriverPropertyBinding("Zone Source Changed", "Occurs when the current source setting for the zone changes.")]
		public IScriptArray ZoneSources
		{
			get
			{
				// If you have an int[] array dedicated to volumes then you could do this:
				return new ScriptArrayMarshalByReference(_zoneSources, new ScriptArraySetInt32Callback(setZoneSource), 1);
			}
		}


		/// <summary>
		/// Expose a Property Array of current Zone Sourcess. All Sources are exposed as thier text source namein the Range 1 to 8
		/// </summary>
		[ScriptObjectPropertyAttribute("Zone Source Names", "Gets the name of the current soruce setting for a zone", "{NAME} Zone, source is {INDEX|1}", null, typeof(ScriptString), 1, MaxSourceCount, "ZoneNames")]
		[SupportsDriverPropertyBinding("Zone Source Names Changed", "Occurs when the current source name setting for the zone changes.")]
		public IScriptArray ZoneSourceNames
		{
			get
			{
				// In the 1st constructor parameter, _zoneNames is of type string[].
				// In the 2nd constructor parameter, 1 is the index of the 1st element in the create ScriptArray.
                return new ScriptArrayMarshalByValue(_zoneSourceNames, 1);
			}
		}


		/// <summary>
		/// Expose a Property Array of the current source names on the Matrix
		/// </summary>
		[ScriptObjectPropertyAttribute("Source Names", "Gets the name of all sources.", "The {NAME} source name for source #{INDEX|1}.", null)]
		public IScriptArray SourceNames
		{
			get
			{
				// In the 1st constructor parameter, _zoneNames is of type string[].
				// In the 2nd constructor parameter, 1 is the index of the 1st element in the create ScriptArray.
                return new ScriptArrayMarshalByValue(_sourceNames, 1);
			}
		}





		/// <summary>
		/// Exposes a method to set the volume on a specified zone in the Matrix
		/// </summary>
		/// <param name="zoneNumber">The Zone Number which we will be Changing the Source on</param>
		/// <param name="volume">The new source we will be applying (Range is 1 to 8)</param>
		[ScriptObjectMethodAttribute("Set Zone Source", "Change the Source for a Zone.", "Set the source for {NAME} zone #{PARAM|0|1} to {PARAM|1|10}.")]
		[ScriptObjectMethodParameterAttribute("zoneNumber", "The zone number.", 1, MaxZoneCount, "ZoneNames")]
		[ScriptObjectMethodParameterAttribute("Source", "Source Input (1 to 8)", 1, MaxSourceCount)]
		public void SetZoneSource(ScriptNumber zoneNumber, ScriptNumber source)
		{
			setZoneSource((int)zoneNumber, (int)source);
		}



		/// <summary>
		/// Exposes a method to Cycle trough each source on a specified zone in the Matrix
		/// </summary>
		/// <param name="zoneNumber">The Zone Number which we will be Changing the Source on</param>
		[ScriptObjectMethodAttribute("Cycle Zone Source", "Cycle the Input Source for a Zone.", "Cycle the Source for {NAME} zone {PARAM|0|1} to next in sequence.")]
		[ScriptObjectMethodParameterAttribute("zoneNumber", "The zone number.", 1, MaxZoneCount, "ZoneNames")]
		public void CycleZoneSource(ScriptNumber zoneNumber)
		{
			// First, ensure that a valid Zone number is provided based on the number of actually configured zones on the matrix

			if ((zoneNumber < 1) || (zoneNumber > _configuredZoneCount))
				throw new ArgumentOutOfRangeException("ZoneNumber", "The zone number must be between 1 and " + _configuredZoneCount);

			// Find out what source number is currenty active for this zone, and increment it
			// then check that this source number is within range, or cycle back to the start

			int nextSource = _zoneSources[(int)zoneNumber];
			nextSource++;

			if ((nextSource < 1) || (nextSource > MaxSourceCount))
				nextSource = 1;

			// Record to the Log that we are about to action a Source Change request
			Logger.InfoFormat("{0}: Zone {1} Set Source to {2}", this.DeviceName, (int)zoneNumber, nextSource);


			// Format the Devices Command String with the paramaters zoneNumber and Source, and send the command.
			//      !{z#}SS{v#}+  
			sendCommand("!" + (int)zoneNumber + "SS" + nextSource + "+");
			sendCommand("!" + (int)zoneNumber + "ZD+");
		}




		#endregion

		#endregion

	}
}