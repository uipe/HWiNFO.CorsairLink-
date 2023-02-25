using System.Timers;
using Topshelf;
using CorsairLink;
using CorsairLink.Synchronization;
using System.Threading;
using Microsoft.Win32;
using System.Collections.Generic;
using System;
using HidSharp;

namespace HWiNFO.CorsairLink
{

    class Program
    {
        static void Main(string[] args)
        {
            HostFactory.Run(windowsService =>
            {
                windowsService.Service<CorsairLinkRegister>(s =>
                {
                    s.ConstructUsing(service => new CorsairLinkRegister());
                    s.WhenStarted(service => service.Start());
                    s.WhenStopped(service => service.Stop());
                });

                windowsService.RunAsLocalSystem();
                windowsService.StartAutomatically();

                windowsService.SetDescription("HWiNFO Corsair Link");
                windowsService.SetDisplayName("HWiNFO Corsair Link");
                windowsService.SetServiceName("HWiNFO.CorsairLink");
            });
        }
    }

    class CorsairLinkRegister
    {
        private System.Timers.Timer timer;
        SupportedDeviceCollection devices;
        Dictionary<string, Dictionary<string, RegistryKey>> regKeys;

        List<IDevice> connectedDevices;


        public CorsairLinkRegister()
        {

            this.timer = new System.Timers.Timer() { AutoReset = true, Interval = 1000 };
            this.timer.Elapsed += ExecuteEvent;

            devices = DeviceManager.GetSupportedDevices(new CorsairDevicesGuardManager(), null);

            connectedDevices = new List<IDevice>();
            regKeys = new Dictionary<string, Dictionary<string, RegistryKey>>();


        }

        private void ExecuteEvent(object sender, ElapsedEventArgs e)
        {

            foreach (var device in connectedDevices)
            {
                Dictionary<string, RegistryKey> deviceKeys;


                if (!regKeys.TryGetValue(device.UniqueId, out deviceKeys))
                {
                    continue;
                }

                device.Refresh();



                foreach (var temp in device.TemperatureSensors)
                {
                    RegistryKey auxKey;

                    if (!deviceKeys.TryGetValue(temp.Name, out auxKey))
                    {
                        continue;
                    }


                    auxKey.SetValue("Name", temp.Name, RegistryValueKind.String);
                    auxKey.SetValue("Value", temp.TemperatureCelsius.HasValue ? temp.TemperatureCelsius : -255, RegistryValueKind.DWord);

                }

                foreach (var speed in device.SpeedSensors)
                {
                    RegistryKey auxKey;

                    if (!deviceKeys.TryGetValue(speed.Name, out auxKey))
                    {
                        continue;
                    }

                    auxKey.SetValue("Name", speed.Name, RegistryValueKind.String);
                    auxKey.SetValue("Value", speed.Rpm.HasValue ? speed.Rpm : 0,RegistryValueKind.DWord);

                }
            }

        }


        public void Start()
        {
            

            connectedDevices = new List<IDevice>();
            regKeys = new Dictionary<string, Dictionary<string, RegistryKey>>();

            foreach (var device in devices)
            {
                if (!device.Connect())
                {
                    Console.WriteLine($"Device '{device.UniqueId}' did not connect!");
                    continue;
                }


                connectedDevices.Add(device);

                Dictionary<string, RegistryKey> auxRegKeys = new Dictionary<string, RegistryKey>();

                regKeys.Add(device.UniqueId, auxRegKeys);

                RegistryKey key = Registry.CurrentUser.CreateSubKey(@$"Software\HWiNFO64\Sensors\Custom\{device.Name}");



                foreach (var temp in device.TemperatureSensors)
                {
                    var auxKey = key.CreateSubKey($"Temp{temp.Channel}");
                    auxKey.SetValue("Name", temp.Name, RegistryValueKind.String);
                    auxKey.SetValue("Value", temp.TemperatureCelsius.HasValue ? temp.TemperatureCelsius : -255, RegistryValueKind.DWord);

                    auxRegKeys.Add(temp.Name, auxKey);
                }

                foreach (var speed in device.SpeedSensors)
                {
                    var auxKey = key.CreateSubKey($"Fan{speed.Channel}");
                    auxKey.SetValue("Name", speed.Name, RegistryValueKind.String);
                    auxKey.SetValue("Value", speed.Rpm.HasValue ? speed.Rpm : 0, RegistryValueKind.DWord);

                    auxRegKeys.Add(speed.Name, auxKey);
                }

            }

            this.timer.Start();
        }

        public void Stop()
        {
            this.timer.Stop();
        }
    }
}