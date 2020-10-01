using FasType.Abbreviations;
using FasType.Converters;
using FasType.Services;
using Microsoft.Extensions.Configuration;
using Serilog;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;

namespace FasType.Storage
{
    public class FileDataStorage : IDataStorage
    {
        readonly string _filepath;
        List<IAbbreviation> _allAbbreviations;
        readonly JsonSerializerOptions serializerOptions;

        ILookup<string, IAbbreviation> AbbreviationsLookup { get; set; }
        IList<IAbbreviation> AllAbbreviations
        {
            get => _allAbbreviations;
            set
            {
                _allAbbreviations = value.ToList();
                AbbreviationsLookup = _allAbbreviations.ToLookup(a => string.Join(null, a.ShortForm.Take(2)), a => a);
            }
        }
        
        public FileDataStorage(IConfiguration _configuration)
        {
            _filepath = _configuration["DataFilePath"];

            serializerOptions = new JsonSerializerOptions();
#if DEBUG
            serializerOptions.WriteIndented = true;
#endif
            serializerOptions.Converters.Add(new IAbbreviationConverter());
            serializerOptions.Converters.Add(new IEnumerableConverter(serializerOptions));

            Task.Run(LoadAsync);
        }

        protected bool Load()
        {
            using var stream = new FileStream(_filepath, FileMode.OpenOrCreate, FileAccess.Read);
            using var reader = new StreamReader(stream);
            string content = reader.ReadToEnd();

            AllAbbreviations = JsonSerializer.Deserialize<IList<IAbbreviation>>(content, serializerOptions);
            Log.Information("Abbreviations Data Storage Loaded.");

            return true;
        }

        protected async Task<bool> LoadAsync()
        {
            using var stream = new FileStream(_filepath, FileMode.OpenOrCreate, FileAccess.Read);

            AllAbbreviations = await JsonSerializer.DeserializeAsync<IList<IAbbreviation>>(stream, serializerOptions);
            Log.Information("Abbreviations Data Storage Loaded.");

            return true;
        }

        protected bool Save()
        {
            using var stream = new FileStream(_filepath, FileMode.OpenOrCreate, FileAccess.Write);
            using var writer = new StreamWriter(stream);
            var ser = JsonSerializer.Serialize(AllAbbreviations, serializerOptions);

            writer.Write(ser);
            
            Log.Information("Abbreviations Data Storage Saved.");

            return true;
        }

        protected async Task<bool> SaveAsync()
        {
            using var stream = new FileStream(_filepath, FileMode.OpenOrCreate, FileAccess.Write);
            await JsonSerializer.SerializeAsync(stream, AllAbbreviations, serializerOptions);

            Log.Information("Abbreviations Data Storage Saved.");

            return true;
        }

        public bool Add(IAbbreviation abbrev)
        {
            AllAbbreviations.Add(abbrev);

            return Save() && Load();
        }

        public async Task<bool> AddAsync(IAbbreviation abbrev)
        {
            AllAbbreviations.Add(abbrev);

            return await SaveAsync() && await LoadAsync();
        }

        public IAbbreviation GetAbbreviation(string shortForm)
        {
            throw new NotImplementedException();
        }

        public IEnumerable<IAbbreviation> GetAbbreviations(string shortForm) => AbbreviationsLookup[string.Join("", shortForm.Take(2))];
    }
}
