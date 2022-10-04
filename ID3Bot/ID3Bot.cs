using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using Dropbox.Api;
using Dropbox.Api.Files;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;

namespace ID3Bot
{
    public class ID3Bot
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

        private MailService _mailService;

        private readonly IConfiguration _configuration;

        public ID3Bot(IConfiguration configuration)
        {
            _configuration = configuration;
            _mailService = new MailService(configuration);
        }

        // At 9:30 every fri
        // 0 30 9 * * Fri

        // Every 20 seconds
        // */30 * * * * *
        [FunctionName("Process")]
        public async Task Run([TimerTrigger("0 30 9 * * Fri")]TimerInfo myTimer, ILogger log)
        {
            log.LogInformation($"I am running!");
            await LoadSongs(log);
            _dropboxClient = InitialiseDropboxClient();

            ListFolderResult files = await _dropboxClient.Files.ListFolderAsync(UntaggedFolderPath);

            log.LogInformation($"Number of files found: {files.Entries.Count}");

            foreach (var file in files.Entries)
            {
                try
                {
                    await DownloadFileAndSaveLocally(file);

                    SetId3Tags(file.Name, log);
                    await UploadFileAndDeleteLocally(file);
                }
                catch (Exception e)
                {
                    log.LogInformation($"Error settings tags for file {file.Name} Exception {e.Message}");
                    continue;
                }
            }

            //await _mailService.SendMailAsync();

            log.LogInformation($"C# Timer trigger function executed at: {DateTime.Now}");
        }

        private DropboxClient InitialiseDropboxClient()
        {            
            string RefreshToken = _configuration["DROPBOX_REFRESH_TOKEN"];
            const string AppKey = "gxy0shh0yvtwjt3";
            string AppSecret = _configuration["DROPBOX_APP_SECRET"];

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

        private static void SetId3Tags(string file, ILogger log)
        {
            var fileName = file.Split("\\")[^1];

            Console.WriteLine(fileName);

            var song = _songs.FirstOrDefault(song => fileName.Contains(song.Title, StringComparison.CurrentCultureIgnoreCase) && song.Artists.Any(fileName.Contains));

            if (song == null)
            {
                Console.WriteLine($"Could not find Google Sheet row for {fileName}");
                log.LogInformation($"Could not find Google Sheet row for {fileName}");
            }

            ValidateSong(song);

            Console.WriteLine($"Song match – {song.Title}");
            log.LogInformation($"Song match – {song.Title}");

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

            log.LogInformation($"ID3 tags added");
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

        private async Task LoadSongs(ILogger log)
        {            
            string credentials = JsonConvert.SerializeObject(new
            {
                type = "service_account",
                project_id = "music-emailer",
                private_key_id = _configuration["GOOGLE_SHEETS_PRIVATE_KEY_ID"],
                private_key = _configuration["GOOGLE_SHEETS_PRIVATE_KEY"],
                client_email = "adam-585@music-emailer.iam.gserviceaccount.com",
                client_id = "108236675381130113762",
                auth_uri = "https://accounts.google.com/o/oauth2/auth",
                token_uri = "https://oauth2.googleapis.com/token",
                auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs",
                client_x509_cert_url = "https://www.googleapis.com/robot/v1/metadata/x509/adam-585%40music-emailer.iam.gserviceaccount.com"
            });

            GoogleCredential credential = GoogleCredential.FromJson(credentials).CreateScoped(Scopes);

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
                log.LogInformation("No data found.");
            }
        }

        private static async Task UploadFileAndDeleteLocally(Metadata file)
        {
            var bytes = await File.ReadAllBytesAsync(file.Name);
            var stream = new MemoryStream(bytes);

            await _dropboxClient.Files.UploadAsync(new UploadArg($"{TaggedFolderPath}/{file.Name}"), stream);

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
