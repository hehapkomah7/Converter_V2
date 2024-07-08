using System;
using System.Xml.Linq;
using System.Text;
using System.IO;
using Microsoft.VisualBasic.FileIO;
using System.Collections.Generic;

namespace ConverterApp
{
    class Program
    {
        static int Main(string[] args)
        {
            if (args.Length != 4)
            {
                Console.WriteLine("4 arguments expected");
                Console.WriteLine("Got arguments: " + args.Length);
                foreach(string arg in args)
                {
                    Console.WriteLine(arg);
                }
                Console.WriteLine("Press 'Enter' to exit the process...");
                while (Console.ReadKey().Key != ConsoleKey.Enter) { }
                return 1;
            }

            var FilePath = args[0];
            var FileOutputPath = args[1];
            var Delimiter = args[2];
            double AdditionalNumber = 0.0;
            if (!double.TryParse(args[3], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out AdditionalNumber))
            {
                Console.WriteLine("Couldn't parse AdditionalNumber:", args[3]);
                Console.WriteLine("Press 'Enter' to exit the process...");
                while (Console.ReadKey().Key != ConsoleKey.Enter) { }
                return 1;
            }

            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            Console.WriteLine($"FilePath = {FilePath}");
            Console.WriteLine($"FileOutputPath = {FileOutputPath}");
            Console.WriteLine($"Delimiter = {Delimiter}");
            Console.WriteLine($"AdditionalNumber = {AdditionalNumber}");
            var program = new Program();
            var converter = new Converter();
            converter.AdditionalNumber = AdditionalNumber;
            var config = program.ReadConfig(FilePath, Delimiter, converter);

            if (config == null)
            {
                Console.WriteLine("Press 'Enter' to exit the process...");
                while (Console.ReadKey().Key != ConsoleKey.Enter) { }
                return 1;
            }
            
            Encoding utf8WithoutBom = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);

            using (
            var sw = new StreamWriter(FileOutputPath, false, utf8WithoutBom)
            )
            {
                sw.Write(config.ToString());
            }

            //Console.WriteLine("Press 'Enter' to exit the process...");
            //while (Console.ReadKey().Key != ConsoleKey.Enter) { }

            return 0;
        }

        XElement MakeConfig(TableParser parser, Converter converter)
        {
            
            var config = new XElement("config");

            var tables = new string[]
            { 
                "vehicles" ,
                "products",
                "tariffs",
                "junctions",
                "shoulders", 
            };

            var methods = new AddTableElement[]
            {
                converter.MakeVehicle,
                converter.MakeProduct,
                converter.MakeTariff,
                converter.MakeJunction,
                converter.MakeShoulder,
            };

            for (int i = 0; i < tables.Length; i++)
            {
                var table = tables[i];
                var method = methods[i];
                var element = converter.MakeGoodTable(parser, "!"+table, table, method);

                if (element == null) return null;

                config.Add(element);
            }

            if (!parser.Tables.ContainsKey("!units"))
            {
                Console.WriteLine("missing table !units :(");
                return null;
            }

            var units = parser.Tables["!units"];

            converter.ParseUnits(units, config);

            return config;

        }

        XElement ReadConfig(string FilePath, string Delimiter, Converter converter)
        {

            using (var sr = new StreamReader(FilePath, Encoding.GetEncoding("windows-1251")))
            {
                using (TextFieldParser readz = new TextFieldParser(sr))
                {
                    readz.TextFieldType = FieldType.Delimited;
                    readz.SetDelimiters(Delimiter);
                    var tableParser = new TableParser();
                    tableParser.Parse(readz);

                    var config = MakeConfig(tableParser, converter);

                    return config;
                }
            }
        }
    }

    public class TableParser
    {

        public Dictionary<string, List<string[]>> Tables = new Dictionary<string, List<string[]>> { };

        public void Parse(TextFieldParser parser)
        {

            var CurrentName = "";
            var CurrentTable = new List<string[]> { };

            while (!parser.EndOfData)
            {
                string[] values = parser.ReadFields();
                if (values[0] == "")
                {
                    continue;
                }

                var name = values[0];
                if (name[0] == '!')
                {
                    CurrentName = name;
                    CurrentTable = new List<string[]> { };
                    Tables.Add(CurrentName, CurrentTable);
                    continue;
                }

                CurrentTable.Add(values);
            }
        }
    }


    public delegate XElement AddTableElement(string[] names, string[] values);
    public class Converter
    {
        public double AdditionalNumber = 0.0;

        public void ParseUnits(List<string[]> table, XElement config)
        {
            var names = table[0];
            var values = table[1];
            var nameConvertion = new Dictionary<string, string>
            {
                {"Объём", "productunit"},
                {"Тариф", "priceunit"},
            };

            for (var i = 0; i < names.Length; i++)
            {
                if (names[i] == "")
                {
                    break;
                }
                config.Add(new XAttribute(nameConvertion[names[i]], values[i]));
            }
        }

        public XElement MakeGoodTable(TableParser parser, string tableName, string elementName, AddTableElement method)
        {
            if (!parser.Tables.ContainsKey(tableName))
            {
                Console.WriteLine($"missing table {tableName} :(");
                return null;
            }

            var lines = parser.Tables[tableName];
            var names = lines[0];

            for (var i = 0; i < names.Length; i++)
            {
                if (names[i] == "")
                {
                    Array.Resize(ref names, i);
                    break;
                }
            }

            var items = new XElement(elementName);

            for (var i = 1; i < lines.Count; i++)
            {
                string[] values = lines[i];

                if (values[0] == "") continue;
                items.Add(method(names, values));
            }

            return items;
        }

       public XElement MakeVehicle(string[] names, string[] values)
        {
            var vehicle = new XElement("vehicle",
                new XAttribute("code", values[0]),
                new XAttribute("name", values[1]),
                new XAttribute("fname", values[2]),
                new XAttribute("dash", values[3])
                );

            return vehicle;
        }

        public XElement MakeProduct(string[] names, string[] values)
        {
            var product = new XElement("product",
                new XAttribute("code", values[0]),
                new XAttribute("name", values[1]),
                new XAttribute("color", values[2])
                );

            for (int i = 3; i < names.Length; i+=3)
            {
                if (values[i] == "") continue;

                var coeffitient = new XElement("coefficient",
                    new XAttribute("vehicle", values[i]),
                    new XAttribute("junctionvalue", values[i + 1]),
                    new XAttribute("shouldervalue", values[i + 2]));
                product.Add(coeffitient);
            }

            return product;
        }

        public XElement MakeTariff(string[] names, string[] values)
        {
            var tariff = new XElement("tariff");
            tariff.Add(new XAttribute("work", values[0]));
            tariff.Add(new XAttribute("vehicle1", values[1]));
            if (values[2] != "")
            {
                tariff.Add(new XAttribute("vehicle2", values[2]));
            }
            tariff.Add(new XAttribute("price", values[3]));

            return tariff;
        }

        XElement MakeCoeffitients(string[] names, string[] values, int start, int end)
        {
            var coefficients = new XElement("coefficients");

            for (int i = start; i < end; i++)
            {
                if (values[i] == "0") continue;

                var item = new XElement("item",
                    new XAttribute("vehicle", names[i]),
                    new XAttribute("value", values[i]));
                coefficients.Add(item);
            }

            return coefficients;
        }

        XElement MakeCapacity(string[] names, string[] values, int start, int end)
        {
            var capacity = new XElement("capacity");

            for (int i = start; i < end; i++)
            {
                if (values[i] == "0"||values[i] == "") continue;

                var item = new XElement("item",
                    new XAttribute("vehicle", names[i]),
                    new XAttribute("value", values[i]));
                capacity.Add(item);
            }

            return capacity;
        }

        void AddProduct(XElement e, string product, string value)
        {
            e.Add(new XElement("item",
                new XAttribute("product", product),
                new XAttribute("value", value)));
        }
        XElement MakeLoading(string[] names, string[] values, int start, int end)
        {
            var loading = new XElement("loading");

            for (int i = start; i < end; i++)
            {
                if (values[i] == "0"||values[i] == "") continue;
                if (!double.TryParse(values[i], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var v))
                {
                    Console.WriteLine("!float.TryParse(values[i], out var v)", values[i]);
                    continue;
                }
                v += AdditionalNumber;
                AddProduct(loading, names[i], v.ToString(System.Globalization.CultureInfo.InvariantCulture));
            }

            return loading;
        }

        public XElement MakeJunction(string[] names, string[] values)
        {
            var coefficients = MakeCoeffitients(names, values, 5, 10);
            var capacity = MakeCapacity(names, values, 10, 15);
            var loading = MakeLoading(names, values, 15, names.Length);

            var junction = new XElement("junction",
                new XAttribute("code", values[0]),
                new XAttribute("name", values[1]),
                new XAttribute("latitude", values[2]),
                new XAttribute("longitude", values[3]),
                new XAttribute("grid", values[4]),
                    coefficients,
                    capacity,
                    loading
                );

            return junction;
        }

        public XElement MakeShoulder(string[] names, string[] values)
        {
            var shoulder = new XElement("shoulder",
                new XAttribute("junction1", values[0]),
                new XAttribute("junction2", values[1]),
                new XAttribute("vehicle", values[2]),
                new XAttribute("capacity", values[3]),
                new XAttribute("distance", values[4]),
                new XAttribute("tariff", values[5])
                );

            return shoulder;
        }
    }
}
