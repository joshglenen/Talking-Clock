# Talking-Clock
Basic c# application using OneClick Installer

To customize:

'Edit "Timer_Tick," remove default cases, and add a case for the time you want. Example below.


        private void myTimer_Tick(object sender, EventArgs e)
        {
            int preSpeech = DateTime.Now.Hour + 1;
            
            if (preSpeech == 0) 
            {
            myTray.ShowBalloonTip(10, "My Notification", "My message", ToolTipIcon.None);
            timer.Interval = 24*MilliSecondsLeftTilTheHour(); 
            return;
            }
            
            timer.Interval = MilliSecondsLeftTilTheHour(); 
            //Note, timer.interval should not normally exceed one day. Max is ~ 24 days.
        } 
        
