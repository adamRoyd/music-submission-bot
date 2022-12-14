using Dropbox.Api;
using Dropbox.Api.Files;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace music_submission_bot
{
    class Program
    {
        const string UntaggedFolderPath = "/[untagged]";
        const string TaggedFolderPath = "/[tagged]";

        static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static readonly string ApplicationName = "Mp3TagBot";
        static readonly string SpreadsheetId = "1_8VQLO2p0rrE4YrYHozMnSDjJG9nki2QoMRF-HUbBok";
        static readonly string sheet = "songs";
        static SheetsService service;
        private static IList<Song> _songs;
        private static DropboxClient _dropboxClient;

        public static async Task Main(string[] args)
        {
            await LoadSongs();
            _dropboxClient = InitialiseDropboxClient();

            ListFolderResult files = await _dropboxClient.Files.ListFolderAsync(UntaggedFolderPath);

            foreach (var file in files.Entries)
            {
                try
                {
                    await DownloadFileAndSaveLocally(file);

                    SetId3Tags(file.Name);
                    await UploadFileAndDeleteLocally(file);
                }
                catch (Exception)
                {
                    // log here
                    continue;
                }
            }
        }

        private static DropboxClient InitialiseDropboxClient()
        {
            //TODO make this a secret
            const string RefreshToken = "pllxHOczeRoAAAAAAAAAAWg1oc1CRI9vr78EuJCdQ6rLdl0REuZnGTZD-uUOMq3L";
            const string AppKey = "gxy0shh0yvtwjt3";
            const string AppSecret = "";

            var httpClient = new HttpClient()
            {
                Timeout = TimeSpan.FromMinutes(20)
            };

            var config = new DropboxClientConfig()
            {
                HttpClient = httpClient
            };

            var client = new DropboxClient(
                RefreshToken,
                AppKey,
                AppSecret,
                config);

            return client;
        }

        private static void SetId3Tags(string file)
        {
            var fileName = file.Split("\\")[^1];

            Console.WriteLine(fileName);

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

        private static async Task UploadFileAndDeleteLocally(Metadata file)
        {
            var bytes = await File.ReadAllBytesAsync(file.Name);
            var stream = new MemoryStream(bytes);

            await _dropboxClient.Files.UploadAsync(new Dropbox.Api.Files.UploadArg($"{TaggedFolderPath}/{file.Name}"), stream);

            File.Delete(file.Name);
        }

        private static async Task DownloadFileAndSaveLocally(Metadata file)
        {
            var downloadResponse = await _dropboxClient.Files.DownloadAsync($"{UntaggedFolderPath}/{file.Name}");
            var content = await downloadResponse.GetContentAsByteArrayAsync();

            await File.WriteAllBytesAsync(file.Name, content);
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
