using System;

namespace GlobalMacroRecorder
{
    [Serializable]
    public class MacroEvent
    {
        public MacroEventType MacroEventType;

        public EventArgs EventArgs;

        public int TimeSinceLastEvent;

        public MacroEvent(MacroEventType macroEventType, EventArgs eventArgs, int timeSinceLastEvent)
        {
            MacroEventType = macroEventType;
            EventArgs = eventArgs;
            TimeSinceLastEvent = timeSinceLastEvent;
        }
    }
}