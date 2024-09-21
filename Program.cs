using System;
using System.IO;
using System.Text;
using HtmlAgilityPack;

class Program
{
    static void Main(string[] args)
    {
        // Check if the filename argument is provided
        if (args.Length == 1)
        {
            string filePath = args[0];

            // Check if the file exists
            if (File.Exists(filePath))
            {
                // Load the HTML document
                HtmlDocument doc = new HtmlDocument();
                doc.Load(filePath);

                // Select elements with class "ts-segment"
                HtmlNodeCollection segmentNodes = doc.DocumentNode.SelectNodes("//div[contains(@class, 'ts-segment')]");

                if (segmentNodes != null)
                {
                    // Create a CSV file with the same name as the HTML file
                    string csvFilePath = Path.ChangeExtension(filePath, "csv");

                    using (StreamWriter writer = new StreamWriter(csvFilePath, false, Encoding.UTF8))
                    {
                        // Write header
                        writer.WriteLine("Name, Timestamp, Text , DurationSeconds, WordCount, SpeakerTotalWordCount, SpeakerTotalDurationSeconds, CurrentWordPerSecond");
                        RowData lastRow = null;
                        
                        //map with SpeakerTotalWordCount SpeakerTotalDurationSeconds  on speaker name
                        Dictionary<string, RowData> speakerData = new Dictionary<string, RowData>();

                        // Process each segment
                        foreach (HtmlNode segmentNode in segmentNodes)
                        {
                            RowData previousRow = lastRow;
                            RowData thisRow = new RowData();
                            
                            // Extract information
                            string name = segmentNode.SelectSingleNode(".//span[@class='ts-name']").InnerText;
                            string timestamp = FormatTimestamp(segmentNode.SelectSingleNode(".//span[@class='ts-timestamp']"));
                            string text = $"\"{DecodeSpecialCharacters(segmentNode.SelectSingleNode(".//span[@class='ts-text']"))}\"";
                            
                            if(name == "" && previousRow != null)
                            {
                                name = previousRow.Name;
                            }
                            else if(name == "")
                            {
                                name = "Unknown";
                            } 
                           
                            
                            thisRow.Timestamp = new SimpleTimeStamp()
                            {
                                Hours = int.Parse(timestamp.Split(".")[0]),
                                Minutes = int.Parse(timestamp.Split(".")[1]),
                                Seconds = int.Parse(timestamp.Split(".")[2])
                            };
                            
                            thisRow.WordCount = text.Split(" ").Length;

                            
                            
                            thisRow.Name = name;
                            thisRow.Text = text;
                            //compute duration for previous row
                            
                            SimpleTimeStamp previousTimeStamp = previousRow?.Timestamp ?? new SimpleTimeStamp()
                            {
                                Hours = 0,
                                Minutes = 0,
                                Seconds = 0
                            };

                            if (previousRow != null)
                            {
                                previousRow.DurationSeconds = Math.Max(previousTimeStamp.DurationToInSeconds(thisRow.Timestamp),1);
                                previousRow.CurrentWordPerSecond =  (double)  previousRow.WordCount / (double)previousRow.DurationSeconds;
                            
                                var preciousSpeakerData = speakerData.GetValueOrDefault(previousRow.Name, null);
                                if(preciousSpeakerData != null)
                                {
                                    preciousSpeakerData.SpeakerTotalDurationSeconds += previousRow.DurationSeconds;
                                    preciousSpeakerData.SpeakerTotalWordCount += previousRow.WordCount;
                                    previousRow.SpeakerTotalDurationSeconds = preciousSpeakerData.SpeakerTotalDurationSeconds;
                                    previousRow.SpeakerTotalWordCount = preciousSpeakerData.SpeakerTotalWordCount;
                                }
                                else
                                {
                                    speakerData.Add(previousRow.Name, new RowData()
                                    {
                                        Name = previousRow.Name,
                                        SpeakerTotalDurationSeconds = previousRow.DurationSeconds,
                                        SpeakerTotalWordCount = previousRow.WordCount
                                    });
                                }
                                writer.WriteLine(previousRow.ToString());
                            }
                            lastRow = thisRow;
                        }
                        
                        writer.WriteLine(lastRow.ToString());
                        
                    }

                    Console.WriteLine($"CSV file written: {csvFilePath}");
                }
                else
                {
                    Console.WriteLine("No segments found with class 'ts-segment'");
                }
            }
            else
            {
                Console.WriteLine("File not found: " + filePath);
            }
        }
        else
        {
            Console.WriteLine("Usage: YourApplicationName filename");
        }
    }

    class SimpleTimeStamp
    {
        public int Hours { get; set; }
        public int Minutes { get; set; }
        public int Seconds { get; set; }
        
        public int ToSeconds()
        {
            return Hours * 3600 + Minutes * 60 + Seconds;
        }
        
        public int DurationToInSeconds( SimpleTimeStamp nextTimeStamp)
        {
            return nextTimeStamp.ToSeconds() - ToSeconds();
        }
    }
    class RowData
    {
        public string Name { get; set; }
        public SimpleTimeStamp Timestamp { get; set; }
        public string Text { get; set; }
        
        //computed
        public int DurationSeconds { get; set; }
        public int WordCount { get; set; }
        public int SpeakerTotalWordCount { get; set; }
        public int SpeakerTotalDurationSeconds { get; set; }
        public double CurrentWordPerSecond { get; set; }

        public override string ToString()
        {
        
            string timestamp = $"{Timestamp.Hours.ToString("00")}:{Timestamp.Minutes.ToString("00")}:{Timestamp.Seconds.ToString("00")}";
            return $"{Name}, {timestamp}, {Text}, {DurationSeconds}, {WordCount}, {SpeakerTotalWordCount}, {SpeakerTotalDurationSeconds}, {CurrentWordPerSecond}";
        
        }
    }
    
    // Format timestamp to a more friendly way for spreadsheets
    static string FormatTimestamp(HtmlNode timestampNode)
    {
        // Extract the timestamp value from the inner text of the anchor tag
        string timestamp = timestampNode.InnerText.Trim();

        // Remove parentheses and convert to a more friendly format
        timestamp = timestamp.Replace("(", "").Replace(")", "").Replace(":", ".");

        return timestamp;
    }

    // Decode special characters in text
    static string DecodeSpecialCharacters(HtmlNode textNode)
    {
        // Decode HTML entities
        string text = System.Net.WebUtility.HtmlDecode(textNode.InnerText.Trim());

        return text;
    }
}
