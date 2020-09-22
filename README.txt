1.  A MUST!
    
    a.  Copy ntsvc.ocx into you system/system32 directory and register it
        using regsvr32

        ex: regsvr32 c:\winnt\system32\ntsvc.ocx

2.  Package and install the program on an NT Machine.  Make sure
    then ntsvc.ocx and the mswinsck.ocx controls are included in the
    installation.  Also include reminder.wav and install it in the
    same directory as broadcast.exe if you want a sound when a message
    comes in.

    Mswinsck.ocx is included with Visual Studio 6, and I suppose other
    Microsoft developing tools.  Distributing and registering this is
    very tricky.  Packaging up mswinsck.ocx with Package in Intallation
    Wizard doesn't register mswinsck.ocx properly.

    I found this out the hard way after installing in on a user's machine.
    Thinking that mswinsck.ocx was properly registere, I got a "can't
    create object" error when the program ran.  Took me a while to narrow it
    down to mswinsck since I had serveral ocxs in this particular program.  

    Anyway, if you have Visual Basic on your machine, you should have no problems
    since the mswinsck.ocx is properly registered when you install it from Visual
    Studio.  If you haven't figured out how to properly distribute mswinsck.ocx
    email me (email below) and I will show you what registry entry is needed to
    properly register mswinsck.

    The reason I didn't include this info in the zip file is I don't know what the
    legal ramifications are with this.

3.  Running the program

    a.  As a standard EXE

        1. Start the broadcast.exe with the -standalone flag on each
           coresponding machines, path may vary of course.

           ex: "c:\program files\..\broadcast.exe" -standalone
    
    b.  As an NT Service

        1. Run broadcast.exe with the -install flag, path may vary
           of course.

           ex: "c:\program files\..\broadcast.exe" -install

           Step 1 will install it as an nt service.

        2. Go into Control Panel/Services, it should be listed 
           as 'SCI Broacasting Service'.  Configure it's 'StarUp' to
           however you want.  It can run under the system account or
           yours or any other account.  You can just leave it to run
           under the system account.  Click 'Automatic' if you want the
           service to start automatically when you start your system. Or
           just start it manually if you want.

        3. To uninstall is as an NT Service, run it with the -uninstall flag.

4.  Sending messages

    a.  Right-click the microphone in the system tray to get the pop-up menu
           
    b.  Click on Send Message

    c.  First thing you need to do is add remote hosts to the "Broadcast to Whom?"
        list.  You DO NOT have to add anyone to test this.  The program is defaulted
        to also send the message to you so you can go to step e if you want. 

        1.  Type in the name of the machine in the 'Host Name' text box
        2.  Type in a friendly name for this machine (an alias), like a person's name.

    d.  There are check boxes next to the names.  The program will broadcast messages
        only to the names that are checked.  This allows control to who you want to
        send to.   

    e.  Type in a message and click send.  It's defaulted to send the message to you
        also so you should see the message.

5.  The Pop-up Menu

    a.  Send Message - allows you to add hosts and send messages
    b.  Send to Self - sends any message you broadcast to yourself also, if this
                       is checked.
    c.  Align Botton - Default.  The message window automatically aligns to the bottom
                       of the screen.
    d.  Align Top    - Aligns the message window to the top. The Align menu items were
                       buggy, so I disabled them
    e.  Loop Messages - Continuously loops through the message(s) if checked.
    f.  Exit - Enabled only in stand-alone mode.  As an NT Service, you stop it using the
               Services Control box in Control Panel.

KNOWN PROBLEM:  When you re-start your system, the Broadcast NT Service starts fine but
the icon is not put in the System Tray so you have to go into control panel/Services
to stop and restart it again.  I'm working on this problem.

Email me at thierry@nv.doe.gov if you have questions.

