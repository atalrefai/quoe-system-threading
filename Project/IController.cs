using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MVCTest
{
    // The Icontroller supports only one functionality that is to increment the value
    public interface IController
    {
        void incvalue();
        void decvalue();
    }

    public class IncController : IController
    {
        IView view;
        IModel model;

        // The controller which implements the IController interface ties the view and model 
        // together and attaches the view via the IModelInterface and addes the event
        // handler to the view_changed function. The view ties the controller to itself.
        public IncController(IView v, IModel m)
        {
            view = v;
            model = m;
            view.setController(this);
            model.attach((IModelObserver)view);
            view.changed += new ViewHandler<IView>(this.view_changed);
        }

        // event which gets fired from the view when the users changes the value
        public void view_changed(IView v, ViewEventArgs e)
        {
            model.setvalue(e.value);
        }

        // This does the actual work of getting the model do the work
        public void incvalue()
        {
            model.increment();
        }
        public void decvalue()
        {
            model.Decrement();
        }
    }
}
