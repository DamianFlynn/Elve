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

namespace DigiNerve
{

	/// <summary>
	/// 
	/// Manuals:
	/// http://forums.sagetv.com/forums/downloads.php?do=file&id=285
	///
	/// Overview:
	/// This driver is used to send Info Popup messages to the SageInfoPopup addin
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
		"SageTV InfoPopup Driver",                                                                      // Display Name 
		"This driver sends messages to the SageTV Infopopup Plug-In. This driver is pretty simple " +   // Description
		"in design and simply allows us to present a popup notification on the SageTV screen. The listener " +
		"should be enabled to accept connections from all servers (0.0.0.0), and must be installed from the " +
		"SageTV 7 plugin repository. The Service provides no feedback so this driver is a single shoot option.",
		"Damian Flynn",                                                                                 // Author
		"Virtual",                                                                                      // Category
		"",                                                                                             // Subcategory
		"SageTVInfoPopup",                                                                              // Default instance name
		DriverCommunicationPort.Network,                                                                // Default Communications Metod
		DriverMultipleInstances.MultiplePerDriverService,                                               // Allow multiple instances
		0,                                                                                              // Major Version
		34,                                                                                             // Minor Version
		DriverReleaseStages.Test,                                                                       // Release Stage
		"SageTV",                                                                                       // Manufacturer Name
		"http://www.sagetv.com",                                                                        // Manufacturer ULR
		null                                                                                            // Registration - for offical use
		)]

	class SageTV_Infopopup_Driver : Driver
	{

		#region Variables

		// Define Communications Protocol Variables
		private TcpCommunication _tcp;
		private string _tcpHostName;
		private int _tcpPort;
		private List<string> _messages;

		private const string TX_TERMINATOR = "";
		private const string RX_TERMINATOR = "\r";

		#endregion


		#region DriverSettingAttribute
		// Provide setting for the driver to be presented in the Elve managment studio

		[DriverSettingAttribute("SageTV IP or Host", "The host name or IP address SageTV Infopopup service is listening on", null, true)]
		public string HostNameSetting
		{
			set { _tcpHostName = value; }
		}


		[DriverSettingAttribute("Port", "The TCP port which the InfoPopup Service is listening on. Defaults to 10629.", 1, 65535, "10629", false)]
		public int PortSetting
		{
			set { _tcpPort = value; }
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
			// the TCP connection on demand, send the message and close the connection again.

			Logger.DebugFormat("{0}: Preparing TCP connection.", this.DeviceName);

			// Open a new TCP session to the target
			_tcp = new TcpCommunication(_tcpHostName, _tcpPort);
			_tcp.Logger = this.Logger;
			_tcp.Delimiter = RX_TERMINATOR;                                     // change the delimiter since it defaults to "\r\n".
			_tcp.CurrentEncoding = System.Text.ASCIIEncoding.ASCII;             // this is unnecessary since it defaults to ASCII
			_tcp.ReceivedDelimitedString += new EventHandler<ReceivedDelimitedStringEventArgs>(_tcp_ReceivedDelimitedString);

			// Set up connection monitor. This is going to be very important to us, as the Infopopup service will not hold
			// a presistant connection, so we are going to use the ConnectionEstablished handler to send any messages we have
			// held in the queue to send.

			_tcp.ConnectionMonitorTimeout = 9000;                              // ensure we receive data at least once every minute so we know the connection is alive
			_tcp.ConnectionMonitorTestRequest = "@message\r";                  // transmit a simple dummy command every 60 seconds to ensure the unit is still connected.
			_tcp.ConnectionEstablished += new EventHandler<EventArgs>(_tcp_ConnectionEstablished);
			_tcp.ConnectionLost += new EventHandler<EventArgs>(_tcp_ConnectionLost);
			_tcp.StartConnectionMonitor();
			
			// Initialize and Clear the queue before we start using it.
			_messages = new List<string> ();
			

			// Since the target service is very dump, and we dont have a lot of possible processing, we can just go ahead
			// and set this drvier into ready to use state.
			return true;

		}

		/// <summary>
		/// Stops the driver by unhooking any event handlers and releasing any used resources.
		/// </summary>
		public override void StopDriver()
		{
			// TODO: Add any necessary cleanup logic here.
			if (_tcp != null)
			{
				_tcp.Dispose();
				_tcp = null;
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


		void _tcp_ConnectionEstablished(object sender, EventArgs e)
		{
			Logger.InfoFormat("{0}: Opened TCP Connection.", this.DeviceName);

			Logger.InfoFormat("{0}: There is {1} messages in queue", this.DeviceName, _messages.Count);
			
			// loop trought the _messages list and send each one in turn
			foreach (string message in _messages)
			{
				// Process the message for transmission
				Logger.InfoFormat("{0}: Sending Message - {1}", this.DeviceName, message);
				_tcp.Send(message + TX_TERMINATOR, 500);
				
				// now that we have processes the message, remove it from the queue
				_messages.Remove(message);
			}
			//IsReady = true;
		}


		void _tcp_ConnectionLost(object sender, EventArgs e)
		{
			Logger.InfoFormat("{0}: Closed TCP Connection.", this.DeviceName);

			// We lost communications, so the driver is no longer in ready status.
			//IsReady = false;
		}


		#endregion

		#region Driver Payload Processing

		/// <summary>
		/// Review the data coming back from the Matrix
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		void _tcp_ReceivedDelimitedString(object sender, ReceivedDelimitedStringEventArgs e)
		{
			try
			{
				string response = e.RawResponse;
				Logger.DebugFormat("{0}: Recived - {1}", this.DeviceName, response);

			}
			catch (Exception ex)
			{
				Logger.ErrorFormat("{0}: Error in the Respone Recived - {1}", this.DeviceName, ex.Message);
			}

		}

		#endregion


		#region Properties and Methods
		
		/// <summary>
		///  Send the command to the matrix
		/// </summary>
		/// <param name="data"></param>
		private void sendCommand(string data)
		{
			Logger.InfoFormat("{0}: Adding Message - {1}", this.DeviceName, data);
			
			_messages.Add(data);
		}


		/// <summary>
		/// Expose a method for Raw Commmand passtrough so that the user has an ability to add additional functionality which is not enabled within the driver.
		/// There is no additional parsing of the input string passed.
		/// </summary>
		/// <param name="rawCommand">String to pass directly to the device</param>
		[ScriptObjectMethod("Display Message", "Display a message on all the extenders", "Display {NAME} message {PARAM|0|Message Text}.")]
		[ScriptObjectMethodParameter("Message", "The Message to display")]
		public void DisplayMessage(ScriptString message)
		{
			sendCommand("@message" + message);
		}

		[ScriptObjectMethod("Display Message And Image", "Display a message and Image on all the extenders", "Display {NAME} message {PARAM|0|Message Text} with image {PARAM|1|image.jpg}.")]
		[ScriptObjectMethodParameter("Message", "The Message to display")]
		[ScriptObjectMethodParameter("Image", "The Image to display")]
		public void DisplayMessage(ScriptString message, ScriptString image)
		{
			sendCommand("@message" + message + "~" + image);
		}

		[ScriptObjectMethod("Display Message And Image on Extender", "Display a message and Image on a specific extender", "Display {NAME} message {PARAM|0|Message Text} with image {PARAM|1|image.jpg} on Extender {PARAM|2|ExtenderUI}.")]
		[ScriptObjectMethodParameter("Message", "The Message to display")]
		[ScriptObjectMethodParameter("Image", "The Image to display")]
		[ScriptObjectMethodParameter("Extender", "The UI Context (MAC Address) of the Extender)")]
		public void DisplayMessage(ScriptString message, ScriptString image, ScriptString uicontext)
		{
			sendCommand("@message" + message + "~" + image + "~" + uicontext);
		}


		// @call{caller_name}~{caller_number}~{description}~{image}~{ui_context}


		[ScriptObjectMethod("Display Caller", "Display a Caller Note on all the extenders", "Display {NAME} message {PARAM|0|Caller Name} and {PARAM|1|Number}.")]
		[ScriptObjectMethodParameter("Caller", "The Callers Name to display")]
		[ScriptObjectMethodParameter("Number", "The Callers Number to display")]
		public void DisplayCaller(ScriptString caller, ScriptString number)
		{
			sendCommand("@call" + caller + "~" + number);
		}

		[ScriptObjectMethod("Display Caller", "Display a Caller Note with Description on all the extenders", "Display {NAME} message {PARAM|0|Caller Name} and {PARAM|1|Number} with {PARAM|2|Description}.")]
		[ScriptObjectMethodParameter("Caller", "The Callers Name to display")]
		[ScriptObjectMethodParameter("Number", "The Callers Number to display")]
		[ScriptObjectMethodParameter("Description", "A Description to display")]
		public void DisplayCaller(ScriptString caller, ScriptString number, ScriptString description)
		{
			sendCommand("@call" + caller + "~" + number + "~" + description);
		}

		[ScriptObjectMethod("Display Caller", "Display a Caller Note with Description and Image on all the extenders", "Display {NAME} message {PARAM|0|Caller Name} and {PARAM|1|Number} with {PARAM|2|Description} using {PARAM|3|image.jpg}.")]
		[ScriptObjectMethodParameter("Caller", "The Callers Name to display")]
		[ScriptObjectMethodParameter("Number", "The Callers Number to display")]
		[ScriptObjectMethodParameter("Description", "A Description to display")]
		[ScriptObjectMethodParameter("Image", "The Image to display")]
		public void DisplayCaller(ScriptString caller, ScriptString number, ScriptString description, ScriptString image)
		{
			sendCommand("@call" + caller + "~" + number + "~" + description + "~" + image);
		}

		[ScriptObjectMethod("Display Caller", "Display a Caller Note with Description and Image on a specific extender", "Display {NAME} message {PARAM|0|Caller Name} and {PARAM|1|Number} with {PARAM|2|Description} using {PARAM|3|image.jpg} on {PARAM|4|extender}.")]
		[ScriptObjectMethodParameter("Caller", "The Callers Name to display")]
		[ScriptObjectMethodParameter("Number", "The Callers Number to display")]
		[ScriptObjectMethodParameter("Description", "A Description to display")]
		[ScriptObjectMethodParameter("Image", "The Image to display")]
		[ScriptObjectMethodParameter("Extender", "The UI Context (MAC Address) of the Extender)")]
		public void DisplayCaller(ScriptString caller, ScriptString number, ScriptString description, ScriptString image, ScriptString uicontext)
		{
			sendCommand("@call" + caller + "~" + number + "~" + description + "~" + image + "~" + uicontext);
		}

		#endregion
		
	}
}
