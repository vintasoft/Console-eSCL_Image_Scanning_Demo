using System;
using System.IO;
using System.Threading;
using Vintasoft.EsclImageScanning;

namespace EsclImageScanningConsoleDemo
{
    public class Program
    {

        static int _currentProgressStep = 0;



        public static int Main(string[] args)
        {
            try
            {
                //EsclEnvironment.EnableDebugging("escl.log");

                // create eSCL device manager
                using (EsclDeviceManager deviceManager = new EsclDeviceManager())
                {
                    // open eSCL device manager
                    deviceManager.Open();

                    Console.WriteLine(string.Format("Searching for eSCL devices in network during {0} seconds... Press 'Enter' key to stop searching.", deviceManager.DeviceSearchTimeout / 1000));
                    DateTime startTime = DateTime.Now;
                    while (DateTime.Now.Subtract(startTime).TotalMilliseconds < deviceManager.DeviceSearchTimeout)
                    {
                        Thread.Sleep(10);
                        if (Console.KeyAvailable)
                        {
                            if (Console.ReadKey().KeyChar == 13)
                            {
                                Console.WriteLine("Searching is canceled.");
                                return -1;
                            }
                        }
                        ShowNextProgressIndicator();
                    }
                    Console.WriteLine();

                    // get count of eSCL devices
                    int deviceCount = deviceManager.Devices.Count;
                    // if eSCL devices are not found
                    if (deviceCount == 0)
                    {
                        Console.WriteLine("Devices are not found.");
                        return -2;
                    }

                    // select eSCL device
                    int deviceIndex = SelectDevice(deviceManager);
                    // if device is not selected
                    if (deviceIndex == 0)
                        return -3;

                    // get selected eSCL device
                    EsclDevice device = deviceManager.Devices[deviceIndex - 1];
                    if (device != null)
                    {
                        // open eSCL device
                        device.Open();

                        // select the scan input source for device
                        SelectDeviceScanInputSource(device);

                        // select the scan intent for eSCL device
                        string scanIntent = SelectEsclScanIntent(device);
                        // if scan intent is selected
                        if (scanIntent != null)
                            // set the scan intent for eSCL device
                            device.ScanIntent = scanIntent;

                        // select the scan color mode for eSCL device
                        EsclScanColorMode? scanColorMode = SelectEsclScanColorMode(device);
                        // if scan color mode is selected
                        if (scanColorMode != null)
                            // set the scan color mode for eSCL device
                            device.ScanColorMode = scanColorMode.Value;

                        // select the scan resolution for eSCL device
                        int? scanXResolution = SelectEsclScanResolution(device);
                        // if scan resolution is selected
                        if (scanXResolution != null)
                        {
                            // set the scan resolution for eSCL device
                            device.ScanXResolution = scanXResolution.Value;
                            device.ScanYResolution = scanXResolution.Value;
                        }

                        EsclScanDocumentFormatExt scanDocumentFormatExt = SelectEsclScanDocumentFormatExt(device);

                        Console.WriteLine("Images acquisition is started...");
                        int imageIndex = 0;
                        // if device should return image in raw format
                        if (scanDocumentFormatExt == EsclScanDocumentFormatExt.OctetStream)
                        {
                            EsclAcquiredImage acquiredImage = null;
                            do
                            {
                                try
                                {
                                    // acquire image from eSCL device
                                    acquiredImage = device.AcquireImageSync();
                                    // if image is acquired
                                    if (acquiredImage != null)
                                    {
                                        Console.WriteLine("Image is acquired.");

                                        string filename = string.Format("scannedImage{0}.jpg", imageIndex);
                                        if (File.Exists(filename))
                                            File.Delete(filename);

                                        // save acquired image to a file
                                        acquiredImage.Save(filename);

                                        Console.WriteLine(string.Format("Image{0} is saved.", imageIndex++));
                                    }
                                    // if image is not acquired
                                    else
                                    {
                                        Console.WriteLine("Scan is completed.");
                                        break;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine(string.Format("Scan is failed: {0}", ex.Message));
                                    break;
                                }
                            }
                            while (acquiredImage != null);
                        }
                        // if device should return image as image file (JPEG or PDF)
                        else
                        {
                            byte[] acquiredImageBytes = null;
                            do
                            {
                                try
                                {
                                    // acquire image bytes from eSCL device
                                    acquiredImageBytes = device.AcquireImageSyncAsFileStream(scanDocumentFormatExt);
                                    // if image is acquired
                                    if (acquiredImageBytes != null)
                                    {
                                        Console.WriteLine("Image is acquired.");

                                        string filename = string.Format("scannedImage{0}", imageIndex);
                                        if (scanDocumentFormatExt == EsclScanDocumentFormatExt.PDF)
                                            filename += ".pdf";
                                        else
                                            filename += ".jpg";
                                        if (File.Exists(filename))
                                            File.Delete(filename);

                                        // save image bytes to a file
                                        File.WriteAllBytes(filename, acquiredImageBytes);

                                        Console.WriteLine(string.Format("Image{0} is saved.", imageIndex++));
                                    }
                                    // if image is not acquired
                                    else
                                    {
                                        Console.WriteLine("Scan is completed.");
                                        break;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine(string.Format("Scan is failed: {0}", ex.Message));
                                    break;
                                }
                            }
                            while (acquiredImageBytes != null);
                        }

                        // close eSCL device
                        device.Close();
                    }

                    // close eSCL device manager
                    deviceManager.Close();
                }

                return 0;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + GetFullExceptionMessage(ex));
                return -1000;
            }
        }

        /// <summary>
        /// Selects eSCL device.
        /// </summary>
        /// <param name="deviceManager">eSCL device manager.</param>
        /// <returns>0 - device is not selected; N - index of selected eSCL device.</returns>
        private static int SelectDevice(EsclDeviceManager deviceManager)
        {
            int deviceCount = deviceManager.Devices.Count;

            Console.WriteLine("Device list:");
            for (int i = 0; i < deviceCount; i++)
            {
                EsclDevice device = deviceManager.Devices[i];
                Console.WriteLine(string.Format("{0}. {1} ({2})", i + 1, device.Name, device.HostUrl));
            }

            int deviceIndex = -1;
            while (deviceIndex < 0 || deviceIndex > deviceCount)
            {
                Console.Write(string.Format("Please select device by entering the device number from '1' to '{0}' or press '0' to cancel: ", deviceCount));
                deviceIndex = Console.ReadKey().KeyChar - '0';
                Console.WriteLine();
            }
            Console.WriteLine();

            return deviceIndex;
        }

        /// <summary>
        /// Selects the scan input source for WIA device.
        /// </summary>
        /// <param name="device">WIA device.</param>
        private static void SelectDeviceScanInputSource(EsclDevice device)
        {
            // if device has flatbed and feeder
            if (device.HasFlatbed && device.HasFeeder)
            {
                if (device.HasDuplex)
                    Console.Write("Device has flatbed and feeder with duplex.");
                else
                    Console.Write("Device has flatbed and feeder without duplex.");

                if (device.IsFeederEnabled)
                    Console.WriteLine(" Now device uses feeder.");
                else
                    Console.WriteLine(" Now device uses flatbed.");


                // ask user to select flatbed or feeder

                int scanInputSourceIndex = -1;
                while (scanInputSourceIndex != 1 && scanInputSourceIndex != 2)
                {
                    Console.Write("What do you want to use: flatbed (press '1') or feeder (press '2'): ");
                    scanInputSourceIndex = Console.ReadKey().KeyChar - '0';
                    Console.WriteLine();
                }
                Console.WriteLine();

                if (scanInputSourceIndex == 1)
                    device.IsFlatbedEnabled = true;
                else
                    device.IsFeederEnabled = true;
            }
            // if device has feeder only
            else if (!device.HasFlatbed && device.HasFeeder)
            {
                if (device.HasDuplex)
                    Console.WriteLine("Device has feeder with duplex.");
                else
                    Console.WriteLine("Device has feeder without duplex.");
                Console.WriteLine();
            }
            // if device has flatbed only
            else if (device.HasFlatbed && !device.HasFeeder)
            {
                Console.WriteLine("Device has flatbed only.");
                Console.WriteLine();
            }
        }

        /// <summary>
        /// Selects the scan intent for eSCL device.
        /// </summary>
        /// <param name="device">eSCL device.</param>
        /// <returns>null - scan intent is not selected; otherwise, selectes scan intent.</returns>
        private static string SelectEsclScanIntent(EsclDevice device)
        {
            string[] supportedScanIntents = device.GetSupportedScanIntents();
            Console.WriteLine("Scan intents:");
            for (int i = 0; i < supportedScanIntents.Length; i++)
            {
                Console.WriteLine(string.Format("{0}. {1}", i + 1, supportedScanIntents[i]));
            }

            int scanIntentIndex = -1;
            while (scanIntentIndex < 0 || scanIntentIndex > supportedScanIntents.Length)
            {
                Console.Write(string.Format("Please select scan intent by entering the number from '1' to '{0}' or press '0' to cancel: ", supportedScanIntents.Length));
                scanIntentIndex = Console.ReadKey().KeyChar - '0';
                Console.WriteLine();
            }
            Console.WriteLine();

            if (scanIntentIndex == 0)
                return null;

            return supportedScanIntents[scanIntentIndex - 1];
        }

        /// <summary>
        /// Selects the scan color mode for eSCL device.
        /// </summary>
        /// <param name="device">eSCL device.</param>
        /// <returns>null - scan color mode is not selected; otherwise, selectes scan color mode.</returns>
        private static EsclScanColorMode? SelectEsclScanColorMode(EsclDevice device)
        {
            EsclScanColorMode[] supportedScanColorModes = device.GetSupportedScanColorModes();
            Console.WriteLine("Scan color modes:");
            for (int i = 0; i < supportedScanColorModes.Length; i++)
            {
                Console.WriteLine(string.Format("{0}. {1}", i + 1, supportedScanColorModes[i]));
            }

            int scanColorModeIndex = -1;
            while (scanColorModeIndex < 0 || scanColorModeIndex > supportedScanColorModes.Length)
            {
                Console.Write(string.Format("Please select scan color mode by entering the number from '1' to '{0}' or press '0' to cancel: ", supportedScanColorModes.Length));
                scanColorModeIndex = Console.ReadKey().KeyChar - '0';
                Console.WriteLine();
            }
            Console.WriteLine();

            if (scanColorModeIndex == 0)
                return null;

            return supportedScanColorModes[scanColorModeIndex - 1];
        }

        /// <summary>
        /// Selects the scan resolution for eSCL device.
        /// </summary>
        /// <param name="device">eSCL device.</param>
        /// <returns>null - scan resolution is not selected; otherwise, selectes scan resolution.</returns>
        private static int? SelectEsclScanResolution(EsclDevice device)
        {
            int[] supportedScanResolutions = device.GetSupportedScanXResolutions();
            Console.WriteLine("Scan resolutions:");
            for (int i = 0; i < supportedScanResolutions.Length; i++)
            {
                Console.WriteLine(string.Format("{0}. {1}", i + 1, supportedScanResolutions[i]));
            }

            int scanResolutionIndex = -1;
            while (scanResolutionIndex < 0 || scanResolutionIndex > supportedScanResolutions.Length)
            {
                Console.Write(string.Format("Please select scan resolution by entering the number from '1' to '{0}' or press '0' to cancel: ", supportedScanResolutions.Length));
                scanResolutionIndex = Console.ReadKey().KeyChar - '0';
                Console.WriteLine();
            }
            Console.WriteLine();

            if (scanResolutionIndex == 0)
                return null;

            return supportedScanResolutions[scanResolutionIndex - 1];
        }

        /// <summary>
        /// Selects the extended scan document format for eSCL device.
        /// </summary>
        /// <param name="device">eSCL device.</param>
        /// <returns>null - extended scan document format is not selected; otherwise, selectes extended scan document format.</returns>
        private static EsclScanDocumentFormatExt SelectEsclScanDocumentFormatExt(EsclDevice device)
        {
            EsclScanDocumentFormatExt[] supportedScanDocumentFormatsExt = device.GetSupportedScanDocumentFormatsExt();
            Console.WriteLine("Scan document formats:");
            for (int i = 0; i < supportedScanDocumentFormatsExt.Length; i++)
            {
                Console.WriteLine(string.Format("{0}. {1}", i + 1, supportedScanDocumentFormatsExt[i]));
            }

            int scanDocumentFormatIndex = -1;
            while (scanDocumentFormatIndex < 1 || scanDocumentFormatIndex > supportedScanDocumentFormatsExt.Length)
            {
                Console.Write(string.Format("Please select scan document format by entering the number from '1' to '{0}': ", supportedScanDocumentFormatsExt.Length));
                scanDocumentFormatIndex = Console.ReadKey().KeyChar - '0';
                Console.WriteLine();
            }
            Console.WriteLine();

            return supportedScanDocumentFormatsExt[scanDocumentFormatIndex - 1];

        }
        /// <summary>
        /// Shows next progress indicator.
        /// </summary>
        private static void ShowNextProgressIndicator()
        {
            string progressStepChars = @"\|/-";
            Console.Write((char)8);
            Console.Write(progressStepChars[_currentProgressStep]);

            _currentProgressStep++;
            if (_currentProgressStep == 4)
                _currentProgressStep = 0;
        }

        /// <summary>
        /// Returns full exception message.
        /// </summary>
        private static string GetFullExceptionMessage(Exception ex)
        {
            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            sb.AppendLine(ex.Message);
            Exception innerException = ex.InnerException;
            while (innerException != null)
            {
                sb.AppendLine(string.Format("Inner exception: {0}", innerException.Message));
                innerException = innerException.InnerException;
            }
            return sb.ToString();
        }

    }
}