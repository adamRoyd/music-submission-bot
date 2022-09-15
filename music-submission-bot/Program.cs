using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace music_submission_bot
{
    class Program
    {
        static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static readonly string ApplicationName = "Mp3TagBot";
        static readonly string SpreadsheetId = "1_8VQLO2p0rrE4YrYHozMnSDjJG9nki2QoMRF-HUbBok";
        static readonly string sheet = "songs";
        static SheetsService service;
        private static IList<Song> _songs;

        public static async Task Main(string[] args)
        {
            await LoadSongs();

            string[] files = Directory.GetFiles(
                args[0],
                "*.mp3",
                SearchOption.AllDirectories);

            foreach (var file in files)
            {
                var fileName = file.Split("\\")[^1];

                Console.WriteLine(fileName);

                // todo match artist too
                // test by downloading the tagged folder and see if it works
                var song = _songs.FirstOrDefault(song => fileName.Contains(song.Title, StringComparison.CurrentCultureIgnoreCase) && song.Artists.Any(fileName.Contains));

                if (song == null)
                {
                    Console.WriteLine($"Could not find Google Sheet row for {fileName}");
                }

                ValidateSong(song);

                Console.WriteLine($"Song match – {song.Title}");

                var track = TagLib.File.Create(file);

                uint bpm;

                uint.TryParse(song.Bpm, out bpm);

                uint year;

                uint.TryParse(song.Year, out year);

                track.Tag.Album = song.Release;
                track.Tag.AlbumArtists = song.Artists;
                track.Tag.BeatsPerMinute = bpm;
                track.Tag.Copyright = $"{song.Year} dojang";
                track.Tag.Comment = song.Comment;
                track.Tag.Composers = song.Artists;
                track.Tag.Genres = new string[] { song.Genre };
                track.Tag.Grouping = song.Grouping;
                track.Tag.Length = track.Properties.Duration.ToString();
                track.Tag.Performers = song.Artists;
                track.Tag.Publisher = song.Publisher;
                track.Tag.Year = year;
                track.Tag.InitialKey = song.Key;

                if (!file.ToLower().Contains("stem"))
                {
                    SetISRCs(file, song, track, fileName);
                    track.Tag.Title = song.Title;
                }

                track.Save();

                Console.WriteLine($"ID3 tags added");
            }
        }

        private static void SetISRCs(string file, Song song, TagLib.File track, string fileName)
        {
            if (song.Versions.Length > 1)
            {
                var songVersion = song.Versions.FirstOrDefault(v => file.Contains(v, StringComparison.CurrentCultureIgnoreCase));
                var versionIndex = Array.FindIndex(song.Versions, s => s == songVersion);

                if (versionIndex != -1)
                {
                    track.Tag.ISRC = song.ISRCs[versionIndex];
                }
                else
                {
                    Console.WriteLine($"Warning - could not find correct version for {fileName}. Please check part of the file name matches a version on the songs spreadsheet exactly");
                }
            }
            else
            {
                track.Tag.ISRC = song.ISRCs[0];
            }
        }

        private static async Task LoadSongs()
        {
            GoogleCredential credential;
            using (var stream = new FileStream("client_secrets.json", FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream)
                    .CreateScoped(Scopes);
            }

            // Create Google Sheets API service.
            service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            var range = $"{sheet}!A:AF";
            SpreadsheetsResource.ValuesResource.GetRequest request =
                service.Spreadsheets.Values.Get(SpreadsheetId, range);

            var response = await request.ExecuteAsync();
            IList<IList<object>> values = response.Values;

            if (values != null && values.Count > 0)
            {
                _songs = new List<Song>();

                var headers = values[0].ToList();

                foreach (var value in values)
                {
                    _songs.Add(new Song
                    {
                        Artists = value[headers.FindIndex(h => h.ToString().Equals("artist(s)", StringComparison.CurrentCultureIgnoreCase))].ToString().Trim().Split(','),
                        Bpm = value[headers.FindIndex(h => h.ToString().Equals("bpm", StringComparison.CurrentCultureIgnoreCase))].ToString(),
                        Comment = value[headers.FindIndex(h => h.ToString().Equals("comment", StringComparison.CurrentCultureIgnoreCase))].ToString(),
                        Genre = value[headers.FindIndex(h => h.ToString().Equals("genre", StringComparison.CurrentCultureIgnoreCase))].ToString(),
                        Grouping = value[headers.FindIndex(h => h.ToString().Equals("grouping", StringComparison.CurrentCultureIgnoreCase))].ToString(),
                        Key = value[headers.FindIndex(h => h.ToString().Equals("key", StringComparison.CurrentCultureIgnoreCase))].ToString(),
                        ISRCs = value[headers.FindIndex(h => h.ToString().Equals("ISRC", StringComparison.CurrentCultureIgnoreCase))].ToString().Trim().Split('\n'),
                        Publisher = "dojang",
                        Release = value[headers.FindIndex(h => h.ToString().Equals("release", StringComparison.CurrentCultureIgnoreCase))].ToString(),
                        Title = value[headers.FindIndex(h => h.ToString().Equals("song"))].ToString(),
                        Year = value[headers.FindIndex(h => h.ToString().Equals("creation year"))].ToString(),
                        Versions = value[headers.FindIndex(h => h.ToString().Equals("versions", StringComparison.CurrentCultureIgnoreCase))].ToString().Trim().Split('\n'),
                    });
                };
            }
            else
            {
                Console.WriteLine("No data found.");
            }
        }

        private static void ValidateSong(Song song)
        {
            if (song.ISRCs.All(a => string.IsNullOrEmpty(a)) || song.ISRCs == null)
            {
                Console.WriteLine($"Warning - ISRCs not found for {song.Title}");
            }

            if (string.IsNullOrEmpty(song.Release))
            {
                Console.WriteLine($"Warning - Release not found for {song.Title}");
            }

            if (song.Artists.All(a => string.IsNullOrEmpty(a)) || song.Artists == null)
            {
                Console.WriteLine($"Warning - Artists not found for {song.Title}");
            }

            if (string.IsNullOrEmpty(song.Bpm))
            {
                Console.WriteLine($"Warning - Bpm not found for {song.Title}");
            }

            if (string.IsNullOrEmpty(song.Year))
            {
                Console.WriteLine($"Warning - Year not found for {song.Title}");
            }

            if (string.IsNullOrEmpty(song.Comment))
            {
                Console.WriteLine($"Warning - Comment not found for {song.Title}");
            }

            if (string.IsNullOrEmpty(song.Grouping))
            {
                Console.WriteLine($"Warning - Grouping not found for {song.Title}");
            }

            if (string.IsNullOrEmpty(song.Genre))
            {
                Console.WriteLine($"Warning - Genre not found for {song.Title}");
            }

            if (string.IsNullOrEmpty(song.Publisher))
            {
                Console.WriteLine($"Warning - Publisher not found for {song.Title}");
            }

            if (string.IsNullOrEmpty(song.Key))
            {
                Console.WriteLine($"Warning - Key not found for {song.Title}");
            }
        }
    }
}
