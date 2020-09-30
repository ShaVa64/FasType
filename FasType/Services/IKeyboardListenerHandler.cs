using FasType.LLKeyboardListener;
using System;
using System.Collections.Generic;
using System.Text;

namespace FasType.Services
{
    public interface IKeyboardListenerHandler
    {
        string CurrentWord { get; }
        //Action<string> CurrentWordCallback { get; }

        //void ListenerEvent(object sender, KeyPressedArgs e);

        void Load(Action<string> currentWordCallback);
        void Close();
    }
}
