using Microsoft.Extensions.DependencyInjection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FasType.Core.Services
{
    public interface IRepositoriesManager
    {
        ILinguisticsRepository Linguistics { get; }
        IAbbreviationsRepository Abbreviations { get; }
        IDictionaryRepository Dictionary { get; }

        void Reload();
    }

    public class RepositoriesManager : IRepositoriesManager
    {
        readonly IServiceProvider _serviceProvider;

        IServiceScope _currentScope;
        ILinguisticsRepository? _linguistics;
        IAbbreviationsRepository? _abbreviations;
        IDictionaryRepository? _dictionary;

        public ILinguisticsRepository Linguistics
        {
            get
            {
                if (_linguistics == null)
                    _linguistics = _currentScope.ServiceProvider.GetRequiredService<ILinguisticsRepository>();
                return _linguistics;
            }
        }
        public IAbbreviationsRepository Abbreviations
        {
            get
            {
                if (_abbreviations == null)
                    _abbreviations = _currentScope.ServiceProvider.GetRequiredService<IAbbreviationsRepository>();
                return _abbreviations;
            }
        }

        public IDictionaryRepository Dictionary
        {
            get
            {
                if (_dictionary == null)
                    _dictionary = _currentScope.ServiceProvider.GetRequiredService<IDictionaryRepository>();
                return _dictionary;
            }
        }

        public RepositoriesManager(IServiceProvider provider)
        {
            _serviceProvider = provider;
            _currentScope = _serviceProvider.CreateScope();
        }

        public void Reload()
        {
            _dictionary?.SaveChanges();
            _linguistics?.SaveChanges();
            _abbreviations?.SaveChanges();

            _currentScope.Dispose();
            _currentScope = _serviceProvider.CreateScope();

            _linguistics = null;
            _abbreviations = null;
            _dictionary = null;
        }
    }
}
