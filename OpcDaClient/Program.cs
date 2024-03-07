using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpcDaClient
{
    class Program
    {
        static void Main(string[] args)
        {
            //Find opc server on local machine
            Opc.IDiscovery discovery = new OpcCom.ServerEnumerator();
            Opc.Server[] serverList = discovery.GetAvailableServers(Opc.Specification.COM_DA_30);
            foreach (Opc.Server item in serverList)
            {
                Console.WriteLine(item.Name);
            }
            //Get opc server
            Opc.Da.Server opcServer = (Opc.Da.Server)serverList[0];
            //Connect server
            opcServer.Connect();

            //Add opc subscription
            Opc.Da.Subscription subscription = null;
            Opc.Da.SubscriptionState subscriptionState = new Opc.Da.SubscriptionState()
            {
                Name = "MySubscription",
                Active = true,
                UpdateRate = 1000,
                Deadband = 0,
                KeepAlive = 1
            };
            subscription = (Opc.Da.Subscription)opcServer.CreateSubscription(subscriptionState);
            subscription.DataChanged += Subscription_DataChanged;
            //Add opc items
            Opc.Da.Item[] items = new Opc.Da.Item[2];
            items[0] = new Opc.Da.Item();
            items[0].ItemName = "Channel1.Device1.TestTag1";

            items[1] = new Opc.Da.Item();
            items[1].ItemName = "Channel1.Device1.TestTag2";
            subscription.AddItems(items);

            //Read opc items
            Opc.Da.ItemValueResult[] readResults = subscription.Read(subscription.Items);
            foreach (var item in readResults)
            {
                Console.WriteLine(item.ItemName + " Value: " + item.Value.ToString());
            }

            //Write opc items
            Opc.Da.ItemValue[] itemValues = new Opc.Da.ItemValue[2];
            itemValues[0] = new Opc.Da.ItemValue();
            itemValues[0].ItemName = "Channel1.Device1.TestTag1";
            itemValues[0].Value = 1;
            itemValues[0].ServerHandle = 1;

            itemValues[1] = new Opc.Da.ItemValue();
            itemValues[1].ItemName = "Channel1.Device1.TestTag2";
            itemValues[1].Value = "MyString";
            itemValues[1].ServerHandle = 2;

            Opc.IdentifiedResult[] writeResults = subscription.Write(itemValues);
            foreach (Opc.IdentifiedResult item in writeResults)
            {
                Console.WriteLine(item.ItemName + ": " + item.ResultID.Name.Name);
            }

            //Browse Items
            Opc.Da.BrowsePosition position = null;
            Opc.Da.BrowseFilters filters = new Opc.Da.BrowseFilters();
            Opc.Da.BrowseElement[] elements = opcServer.Browse(null, filters, out position);
            foreach (var item in elements)
            {
                if (item.HasChildren)
                {
                    Opc.ItemIdentifier id = new Opc.ItemIdentifier(item.Name);
                    Opc.Da.BrowseElement[] children = opcServer.Browse(id, filters, out position);
                }
            }
            //Browse all items
            BrowsAllElement("", ref opcServer);


            //Clean up connection
            //remove items
            subscription.RemoveItems(subscription.Items);
            opcServer.CancelSubscription(subscription);
            opcServer.Disconnect();
            opcServer.Dispose();
            subscription.Dispose();
            Console.ReadLine();
        }

        private static void BrowsAllElement(string itemName, ref Opc.Da.Server opcServer)
        {
            Opc.Da.BrowseFilters filters = new Opc.Da.BrowseFilters();
            Opc.Da.BrowsePosition position = null;
            Opc.ItemIdentifier id = new Opc.ItemIdentifier(itemName);
            Opc.Da.BrowseElement[] children = opcServer.Browse(id, filters, out position);
            if (children != null)
            {
                foreach (var item in children)
                {
                    if (item.HasChildren)
                    {
                        if (string.IsNullOrEmpty(itemName))
                        {
                            Console.WriteLine(item.Name);
                            BrowsAllElement(item.Name, ref opcServer);

                        }
                        else
                        {
                            Console.WriteLine(itemName + "." + item.Name);
                            BrowsAllElement(itemName + "." + item.Name, ref opcServer);
                        }
                    }
                    else
                    {
                        //The item you want
                        Console.WriteLine(itemName + "." + item.Name);
                    }
                }
            }
        }

        private static void Subscription_DataChanged(object subscriptionHandle, object requestHandle, Opc.Da.ItemValueResult[] values)
        {
            //Do not recommend using this event, using OPC AE Client instead.
        }
    }
}
