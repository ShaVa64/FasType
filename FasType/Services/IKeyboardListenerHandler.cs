using FasType.LLKeyboardListener;
using System;
using System.Collections.Generic;
using System.Text;

namespace FasType.Services
{
    public interface IKeyboardListenerHandler
    {
        string CurrentWord { get; }

        void Load();
        void Close();

        void Pause();
        void Continue();
    }
}