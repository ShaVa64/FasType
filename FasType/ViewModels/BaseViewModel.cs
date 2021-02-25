using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using FasType.Models;

namespace FasType.ViewModels
{
    public class BaseViewModel : ObservableObject
    {
        readonly WeakReference<Window> windowWeakRef;

        protected Window Window 
        {
            get 
            {
                if (windowWeakRef.TryGetTarget(out Window w))
                    return w;
                return null;
            }
        }

        public BaseViewModel(Window w)
        {
            windowWeakRef = new(w);
        }
    }
}
