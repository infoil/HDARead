﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using NDesk.Options;

using Opc;
using OpcCom;

using System.Diagnostics;
using Microsoft.VisualBasic.Logging;

using System.IO;

/* TODO:
 * +aggregate enum
 * +resample interval
 * +file output
 * exceptions
 * +read raw
 * +input file with tag list
 * why we use OPCTrend and not OPCServer read method
 * maybe its better to query data tag by tag (because 'tag not found' exception doesn't show which tag is wrong)
 *    or maybe validate tags before read and delete them from query?
 * command line parameter to display output quality or not
 * command line parameter to 'merged' output format (single timestamp column for all tags)
 * OutputTable: What if different tags have different number of points?!
 */

namespace HDARead {

    enum eOutputFormat {
        LIST = 1,
        TABLE = 2,
        MERGED = 3
    }

    class Program {
        static string Host = null;
        static string Server = null;
        static string StartTime = "NOW-1H";
        static string EndTime = "NOW";
        static int Aggregate = (int)HDAClient.OPCHDA_AGGREGATE.AVERAGE;
        static string OutputTimestampFormat = null;
        static int MaxValues = 10;
        static int ResampleInterval = 0;
        static string OutputFileName = null;
        static string InputFileName = null;
        static bool ReadRaw = false;
        static bool Help = false;
        static eOutputFormat OutputFormat = eOutputFormat.LIST;
        static List<string> Tagnames = new List<string>();
        static string OptionDescription = null;

        static TraceSource _trace = new TraceSource("ConsoleApplicationTraceSource");

        static void Main(string[] args) {

            if (!ParseCommandLine(args)) {
                return;
            }
            if (!CheckOptions()) {
                return;
            }
            ShowInfo();

            Opc.Hda.ItemValueCollection[] OPCHDAItemValues = null;
            try {
                bool res = false;
                var srv = new HDAClient();
                _trace.TraceEvent(TraceEventType.Verbose, 0, "Created HDAClient");
                if (srv.Connect(Host, Server)) {
                    _trace.TraceEvent(TraceEventType.Verbose, 0, "Connected. Going to read.");
                    res = srv.Read(StartTime, EndTime, Tagnames.ToArray(), Aggregate, MaxValues, ResampleInterval, ReadRaw, out OPCHDAItemValues);
                } else {
                    Console.WriteLine("HDARead unable to connect to OPC server.");
                }
                srv.Disconnect();
                if (!res) {
                    Console.WriteLine("HDARead Error.");
                    return;
                }
            } catch (Exception e) {
                Console.WriteLine(e.Message);
                return;
            }
            if (OPCHDAItemValues == null) {
                Console.WriteLine("HDARead returned null.");
                return;
            } else {
                Console.WriteLine("HDARead OK.");
            }
            _trace.TraceEvent(TraceEventType.Verbose, 0, "Number of tags = OPCHDAItemValues.Count()={0}", OPCHDAItemValues.Count());

            StreamWriter writer;
            if (!string.IsNullOrEmpty(OutputFileName)) {
                writer = new StreamWriter(OutputFileName);
            } else {
                writer = new StreamWriter(Console.OpenStandardOutput());
                writer.AutoFlush = true;
                Console.SetOut(writer);
            }

            switch (OutputFormat) {
                case eOutputFormat.LIST:
                    OutputList(writer, OPCHDAItemValues, OutputTimestampFormat);
                    break;
                case eOutputFormat.MERGED:
                    OutputMerged(writer, Merge(OPCHDAItemValues), OutputTimestampFormat);
                    break;
                case eOutputFormat.TABLE:
                    OutputTable(writer, OPCHDAItemValues, OutputTimestampFormat);
                    break;
            }

            if (!string.IsNullOrEmpty(OutputFileName)) {
                writer.Close();
                Console.WriteLine("Data were written to file {0}.", OutputFileName);
            }
            
            return;
        }

        static bool ParseCommandLine(string[] args) {
             var p = new OptionSet() {
   	            { "n=|node=",               "Remote computer name (optional)",  v => Host = v },
   	            { "s=|server=",             "OPC HDA server name (required)",   v => Server = v },
   	            { "from=|start=|begin=",    "Start time (abs. or relative), default NOW-1H",    
                                                                                v => StartTime = v ?? "NOW-1H"},
   	            { "to=|end=",               "End time (abs. or relative), default NOW",      
                                                                                v => EndTime = v ?? "NOW"},
   	            { "a=|agg=",                "Aggregate (see spec)",             v => Aggregate = GetHDAAggregate(v) },
                { "r=|resample=",           "Resample interval (in seconds), 0 - return just one value (see OPC HDA spec.)",  
                                                                                v => ResampleInterval = Int32.Parse(v)},
                { "raw",                    "Read raw data (if omitted, read processed data) ",  
                                                                                v => ReadRaw = v != null},
                { "m=|maxvalues=",          "Maximum number of values to load (should be checked at OPC server side, but doesn't work)", 
                                                                                v => MaxValues = Int32.Parse(v)},
                { "t=|tsformat=",           "Output timestamp format to use",   v => OutputTimestampFormat = v},
                { "f=",                     "Output format (LIST or TABLE or MERGED)",   
                                                                                v => OutputFormat = GetOutputFormat(v)},
                { "o=|output=",             "Output filename (if omitted, output to console)",   
                                                                                v => OutputFileName = v},
                { "i=|input=",              "Input filename with list of tags (if omitted, tag list must be provided as command line argument)",   
                                                                                v => InputFileName = v},
                { "h|?|help",               "Show help",                        v => Help = v != null},
                { "<>",                     "List of tag names",                v => Tagnames.Add (v)},
            };

            Console.Write("HDARead: ");
            try {
                p.Parse(args);
            } catch (OptionException e) {
                Console.WriteLine(e.Message);
                Console.WriteLine("Try `HDARead --help' for more information.");
                return false;
            }

            StringBuilder sb = new StringBuilder();
            TextWriter tw = new StringWriter(sb);
            if(tw != null) {
                p.WriteOptionDescriptions(tw);
                tw.Flush();
                OptionDescription = sb.ToString();
            }
            return true;
        }


        static bool CheckOptions() {
            if (Help) {
                ShowHelp();
                return false;
            }

            if (string.IsNullOrEmpty(Server)) {
                Console.WriteLine("Missing required option s=|server=");
                return false;
            }

            if (string.IsNullOrEmpty(InputFileName)) {
                if (Tagnames.Count() < 1) {
                    Console.WriteLine("No tagnames were specified.");
                    return false;
                }
            } else {
                if (Tagnames.Count() > 0) {
                    Console.WriteLine("If the input file is specified, no tags may be entered as command line argument");
                    return false;
                }
                // try catch !!!
                Tagnames = File.ReadLines(InputFileName).ToList();
                if (Tagnames.Count() < 1) {
                    Console.WriteLine("No tagnames were specified.");
                    return false;
                }
            }
            return true;
        }

        static void ShowInfo() {
            Console.WriteLine("HDARead is going to:");
            Console.WriteLine("\t connect to OPC HDA server named '{0}' on computer '{1}'", Server, Host);
            Console.WriteLine("\t and read {0} data for the period from '{1}' to '{2}'", ReadRaw ? "raw" : "processed", StartTime, EndTime);
            Console.WriteLine("\t with resample interval {0} seconds", ResampleInterval, EndTime);
            Console.WriteLine("\t for the following tags:");
            foreach (string t in Tagnames) {
                Console.WriteLine("\t\t" + t);
            }
            Console.WriteLine("\t No more than {0} values should be loaded (checked only at OPC server side).", MaxValues);
        }

        static void OutputList(StreamWriter sw, Opc.Hda.ItemValueCollection[] OPCHDAItemValues, string OutputTimestampFormat) {
            string ts;
            for (int i = 0; i < OPCHDAItemValues.Count(); i++) {
                sw.WriteLine();
                sw.WriteLine("\tTag ({0} of {1}): {2} ({3} values):", i + 1, OPCHDAItemValues.Count(), OPCHDAItemValues[i].ItemName, OPCHDAItemValues[i].Count);
                sw.WriteLine("{0,20}{1,20}{2,20}", "Timestamp", "Value", "Quality");

                for (int j = 0; j < OPCHDAItemValues[i].Count; j++) {
                    if (OutputTimestampFormat != null) {
                        ts = OPCHDAItemValues[i][j].Timestamp.ToString();
                    } else {
                        ts = OPCHDAItemValues[i][j].Timestamp.ToString(OutputTimestampFormat);
                    }
                    sw.WriteLine("{0,20}{1,20}{2,20}", ts, OPCHDAItemValues[i][j].Value.ToString(), OPCHDAItemValues[i][j].Quality.ToString());
                }
            }
        }

        static void OutputTable(StreamWriter sw, Opc.Hda.ItemValueCollection[] OPCHDAItemValues, string OutputTimestampFormat) {
            // header
            for (int i = 0; i < OPCHDAItemValues.Count(); i++) {
                if (i > 0) {
                    sw.Write(",");
                }
                sw.Write(",{0} timestamp, {0} value, {0} quality", OPCHDAItemValues[i].ItemName);
            }
            sw.WriteLine();
            string ts;

            // What if different tags have different number of points?!
            int max_rows = OPCHDAItemValues[0].Count;
            for (int i = 1; i < OPCHDAItemValues.Count(); i++) {
                if (max_rows < OPCHDAItemValues[0].Count) {
                    max_rows = OPCHDAItemValues[0].Count;
                }
            }

            for (int j = 0; j < max_rows; j++) {
                for (int i = 0; i < OPCHDAItemValues.Count(); i++) {
                    if (i > 0) {
                        sw.Write(",");
                    }
                    if (j < OPCHDAItemValues[i].Count) {
                        ts = GetDatetimeStr(OPCHDAItemValues[i][j].Timestamp, OutputTimestampFormat);
                        sw.Write("{0},{1},{2}", ts, OPCHDAItemValues[i][j].Value.ToString(), OPCHDAItemValues[i][j].Quality.ToString());
                    } else {
                        sw.Write(",,");
                    }
                }
                sw.WriteLine();
            }
        }
        static void OutputMerged(StreamWriter sw, Opc.Hda.ItemValueCollection[] OPCHDAItemValues, string OutputTimestampFormat) {
            // header
            sw.Write("Timestamp, {0} value, {0} quality", OPCHDAItemValues[0].ItemName);
            for (int i = 1; i < OPCHDAItemValues.Count(); i++) {
                sw.Write(",{0} value, {0} quality", OPCHDAItemValues[i].ItemName);
            }
            sw.WriteLine();
            string ts;
            for (int j = 0; j < OPCHDAItemValues[0].Count; j++) {
                ts = GetDatetimeStr(OPCHDAItemValues[0][j].Timestamp, OutputTimestampFormat);
                sw.Write("{0}", ts);
                for (int i = 0; i < OPCHDAItemValues.Count(); i++) {
                    if (OPCHDAItemValues[i][j].Value == null) {
                        sw.Write(",");
                    } else {
                        sw.Write(",{0}", OPCHDAItemValues[i][j].Value.ToString());
                    }
                    if (OPCHDAItemValues[i][j].Quality == null) {
                        sw.Write(",");
                    } else {
                        sw.Write(",{0}", OPCHDAItemValues[i][j].Quality.ToString());
                    }
                }
                sw.WriteLine();
            }
        }

        static  Opc.Hda.ItemValueCollection[] Merge(Opc.Hda.ItemValueCollection[] OPCHDAItemValues) {
            int n_tags = OPCHDAItemValues.Count();
            _trace.TraceEvent(TraceEventType.Verbose, 0, "Starting merge. n_tags = {0}", n_tags);
            var MergedValues = new Opc.Hda.ItemValueCollection[n_tags];
            for (int i = 0; i < n_tags; i++) {
                MergedValues[i] = new Opc.Hda.ItemValueCollection(new Opc.ItemIdentifier(OPCHDAItemValues[i]));
            }

            // init pointer (row numbers) for each column
            int[] row = new int[n_tags];
            for (int i = 0; i < n_tags; i++) {
                row[i] = 0;
            }

            bool have_more_data = true;
            while (have_more_data) {

                string msg = "rows: ";
                for (int i = 0; i < n_tags; i++) {
                    msg += row[i] + ", ";
                }
                _trace.TraceEvent(TraceEventType.Verbose, 0, msg);

                // find min timestamp
                int min_ts_col = -1;
                DateTime min_ts = System.DateTime.MaxValue;
                _trace.TraceEvent(TraceEventType.Verbose, 0, "Looking for min timestamp");

                for (int i = 0; i < n_tags; i++) {
                    if (row[i] >= OPCHDAItemValues[i].Count) {
                        _trace.TraceEvent(TraceEventType.Verbose, 0, "Check tag {0}: no data", OPCHDAItemValues[i].ItemName);
                    } else {
                        _trace.TraceEvent(TraceEventType.Verbose, 0, "Check tag {0}: {1}", OPCHDAItemValues[i].ItemName, OPCHDAItemValues[i][row[i]].Timestamp);
                        if (min_ts > OPCHDAItemValues[i][row[i]].Timestamp) {
                            min_ts = OPCHDAItemValues[i][row[i]].Timestamp;
                            min_ts_col = i;
                        }
                    }
                }
                _trace.TraceEvent(TraceEventType.Verbose, 0, "Min timestamp = {0}, index = {1}", min_ts, min_ts_col);
                have_more_data = false;
                // copy value with this timestamp to output array
                for (int i = 0; i < n_tags; i++) {
                    if ((row[i] < OPCHDAItemValues[i].Count) && (OPCHDAItemValues[i][row[i]].Timestamp.Equals(min_ts))) {
                        MergedValues[i].Add(OPCHDAItemValues[i][row[i]]);
                        _trace.TraceEvent(TraceEventType.Verbose, 0, "Copying: {0}, {1}, {2}",
                            OPCHDAItemValues[i].ItemName,
                            OPCHDAItemValues[i][row[i]].Timestamp.ToString(),
                            OPCHDAItemValues[i][row[i]].Value.ToString());
                        row[i]++;
                        if (row[i] < OPCHDAItemValues[i].Count) {
                            have_more_data = true;
                        }
                    } else {
                        // if there is no value for this timestamp, fill blank
                        var itemvalue = new Opc.Hda.ItemValue();
                        itemvalue.Timestamp = min_ts;
                        itemvalue.Value = null;
                        var q = new Opc.Da.Quality();
                        q.QualityBits = Opc.Da.qualityBits.uncertain;
                        itemvalue.Quality = q;
                        itemvalue.HistorianQuality = Opc.Hda.Quality.NoData;
                        MergedValues[i].Add(itemvalue);

                        _trace.TraceEvent(TraceEventType.Verbose, 0, "Filling with blank: {0}, {1}",
                            OPCHDAItemValues[i].ItemName,
                            itemvalue.Timestamp.ToString());
                    }
                }
            }
            return MergedValues;
        }
        static void ShowHelp() {
            Console.WriteLine("Usage: HDARead OPTIONS tag1 tag2 tag3");
            Console.WriteLine("HDARead is used to read the data from OPC HDA server.");
            Console.WriteLine();
            Console.WriteLine("Options:");
            Console.WriteLine(OptionDescription);
            Console.WriteLine();
            Console.WriteLine("Aggregates:");
            foreach (string agg in Enum.GetNames(typeof(HDAClient.OPCHDA_AGGREGATE))) {
                Console.Write(" " + agg);
            }
                
        }

        static int GetHDAAggregate(string str) {
            HDAClient.OPCHDA_AGGREGATE Value;

            if (Enum.TryParse(str, out Value) && Enum.IsDefined(typeof(HDAClient.OPCHDA_AGGREGATE), Value))
                return (int)Value;
            else
                throw new NDesk.Options.OptionException("Wrong aggregate: " + str, "-a");
        }

        static eOutputFormat GetOutputFormat(string str) {
            eOutputFormat Value;
            if (string.IsNullOrEmpty(str))
                return eOutputFormat.LIST;

            if (Enum.TryParse(str, out Value) && Enum.IsDefined(typeof(eOutputFormat), Value))
                return Value;
            else
                throw new NDesk.Options.OptionException("Wrong output format:: " + str, "-f");
        }

        static string GetDatetimeStr(DateTime dt, string fmt) {
            string ts;
            if (string.IsNullOrEmpty(fmt)) {
                ts = dt.ToString();
            } else {
                ts = dt.ToString(fmt);
            }
            return ts;
        }
    }
}
