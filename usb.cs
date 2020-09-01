using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Windows.Devices.Enumeration;
using Windows.Foundation;

namespace UsbDeviceRC
{
    [DataContract]
    internal sealed class UsbDeviceForJson
    {
        [DataMember]
        public string deviceID { get; set; }
        [DataMember]
        public string deviceName { get; set; }
        [DataMember]
        public string deviceProtocol { get; set; }
        [DataMember]
        public string productId { get; set; }
        [DataMember]
        public string vendorId { get; set; }
        [DataMember]
        public string deviceClass { get; set; }
        [DataMember]
        public string deviceSubclass { get; set; }
        [DataMember]
        public string interfaceCount { get; set; }
    }
    /// <summary>json化用クラス</summary>
    [DataContract]
    internal sealed class UsbDeviceListForJson
    {
        [DataMember]
        public IEnumerable<UsbDeviceForJson> usabledeviceinfo { get; set; }
    }

    public static class UsbDeviceRC
    {
        /// <summary>USBデバイスの一覧を返す</summary>
        /// <returns>USBデバイスの一覧を含むJSON文字列</returns>
        public static IAsyncOperation<string> GetUSBConnectedListAsync() =>
           GetUSBConnectedListAsync_().AsAsyncOperation();

        private const string PATTERN =
            @"^\\\\\?\\USB#VID_(?<vid>[0-9A-F]{4})&PID_(?<pid>[0-9A-F]{4})";

        private static Regex regUsbDevices = new Regex(PATTERN);

        private static async Task<string> GetUSBConnectedListAsync_()
        {
            Debug.WriteLine("UsbDeviceRC.GetUSBConnectedListAsync_ - 01");
            try
            {
                var selector =
                    "System.Devices.InterfaceEnabled:=System.StructuredQueryType.Boolean#True";

                // DeviceInformationの一覧を取得
                var deviceInformationCollection =
                    await DeviceInformation.FindAllAsync(selector);


                Func<String, String> HexToDec =
                    hex => Convert.ToInt32(hex, 16).ToString();

                var usbDeviceForJsons = new List<UsbDeviceForJson>();
                var isUsbDeviceExist = false;
                var dicPidVid = new Dictionary<string, string>();
                foreach (var device in deviceInformationCollection)
                {
                    if (device.Id.StartsWith(@"\\?\USB#"))
                    {
                        var match = regUsbDevices.Match(device.Id);
                        if (string.IsNullOrEmpty(match.Value))
                        {
                            continue;
                        }
                        var pid = HexToDec(match.Groups["pid"].Value);
                        var vid = HexToDec(match.Groups["vid"].Value);

                        isUsbDeviceExist = false;

                        if (dicPidVid.ContainsKey(pid) && vid.Equals(dicPidVid[pid]))
                        {
                            isUsbDeviceExist = true;
                        }
                        else
                        {
                            dicPidVid.Add(pid, vid);
                        }

                        if (isUsbDeviceExist == false)
                        {
                            var usbDevice = new UsbDeviceForJson();
                            usbDevice.deviceName = device.Name;
                            usbDevice.deviceID = device.Id;
                            usbDevice.productId = pid;
                            usbDevice.vendorId = vid;

                            usbDeviceForJsons.Add(usbDevice);
                        }
                    }
                }

                // json化用クラスUsbDeviceListのインスタンスを生成
                var usbDeviceListForJson =
                   new UsbDeviceListForJson() { usabledeviceinfo = usbDeviceForJsons };

                // JSON文字列を生成する
                var jsonString = usbDeviceListForJson.ToJsonString();

                return
                    usbDeviceForJsons.Any() ? jsonString : "{}";
            }
            catch
            {
                return "{}";
            }
            finally
            {
                Debug.WriteLine("UsbDeviceRC.GetUSBConnectedListAsync_ - 02");
            }
        }

        /// <summary>与えられたオブジェクトよりJSON文字列を生成する</summary>
        /// <param name="obj">JSON文字列化を行うオブジェクト</param>
        /// <returns>JSON文字列</returns>
        private static string ToJsonString<T>(this T obj)
        {
            Debug.WriteLine("UsbDeviceRC.ToJsonString - 01");
            try
            {
                using (var sr = new StreamReader(new MemoryStream()))
                {
                    new DataContractJsonSerializer(typeof(T)).WriteObject(sr.BaseStream, obj);
                    sr.BaseStream.Position = 0;
                    return sr.ReadToEnd();
                }
            }
            finally
            {
                Debug.WriteLine("UsbDeviceRC.ToJsonString - 02");
            }
        }
    }
}
