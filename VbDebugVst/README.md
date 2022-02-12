# VbDebugVst
## VB6 VST Debugger Utility

Hello everyone!

This utility is intended for debug your VST plugins written using **VbVst** framework. 
Because you can't debug the native-dlls in VB6 IDE (and VST plugins are the native-dlls) i developed this very basic utility.
After you run this utility you need to configure **ProgID** of your VB6 project. 
You should go to **Debug->Setup** menu option and type the object's **ProgID** (\[ProjectName.ClassName\]). 
You can enable/disable the editor if you wish. The editor is the custom GUI which your plugin can provide. 
VST supports the editorless plugins as well the host uses its own controls to represents the parameters.
This settings are saved to config.ini file so they are loaded during program initialization.

After configuring utility you can create plugin instances. Just run your plugin project (select *Wait for components to be created* option) and press **Create plugin** button.
The utility searches for the needed VB6 instance, injects to it and creates your plugin object. 
You can press the stop button, pause code, do step debugging etc as if you debug an usual ActiveX dll. You can open the log window to see the errors or the plugin properties.

You can load an audio file (WAV file) and use it in the simple **Song Editor**. The Song Editor allows you to place the audio to 8 bars and play them through your VST plugin (1 track). 
For example you can load a drum loop and propagate it through the song. The repository contains the example loop. 
The important thing is the utility supports only the basic formats 8 or 16 bits WAVE-PCM audio.

Also it allows to place the MIDI pattern with the notes in this 8 bars (2 track). VST plugin will get the note on/note off events according the pattern. 
You can edit the pattern in the simple **Pattern Editor**. Just place the notes using the left mouse button and delete them by the right button. 
You can hold Ctrl-key to change the note velocity. Using Pattern Editor you can test your **ProcessEvents** method.

The last track is the event editor. This is the very basic editor for debugging purposes. You can edit the automation here. Richt-click on the panel allows you to select the parameter.
The first 5 items are related to MIDI events like pitch bend/modulation wheel etc. Other items are the plugin parameters. You can't automate a non-automable parameter.
This track allows you also debug your **ProcessEvents** method and check the parameters automation. The recording isn't supported. 
The right click context menu allows you to clear the track.

You can save and load the plugin state under the File menu. This allows you to check your **SetStateChunk/GetStateChunk** methods. 
The important thing is the state is saved with the plugin unique id. So if you have changed the ID the state won't load anymore. 
If plugin doesn't support the persistence the utility just saves all the parameters.

Under **Debug->Options** you can change the current **Tempo**, the **Sample Rate** and the **BlockSize**. 
Because the utility is very basic the audio resampling uses the very straight algorithm without any filtering. 
Also if you change the sample rate to a lower value the audio track saves this and if you then restore it back the audio can't restore it because the information is lost.

When you don't use the VST-editor feature (custom UI) the utility creates internal window with the plugin parameters. You can work with this MDI-child window as usual. 
Another case if you use custom UI. The utility don't place the editor UI inside a MDI child window but on an external window. 
This is because the utility is intended for debugging and your code can be paused. 
If your editor window lays on a freezed window (because utility waits for response) the input queue is freezed too. So you can't use mouse/keyboard input. 
The utility creates the custom UI in the plugin process therefore such problems are solved. 
The professional DAW can place windows between processes because there is no pausing like in debugging.

## How does it work?

The utility uses a special native-dll written in VB6 (dll directory in this repository). This dll shares the needed data (like samples/events) between 2 processes using file-mapping.
When you create a plugin by its ProgID the dll searches for debugging VB6.exe instance using some undocumented features which i found out using reverse engineering.
The VB6 executable saves the COM-Thread-ID ([CoGetCurrentProcess](https://docs.microsoft.com/en-us/windows/win32/api/combaseapi/nf-combaseapi-cogetcurrentprocess)) along with project object address into registry.
The code scans all the entries in registry with the corresponding ProgIds. Then it scans for the VB6 IDE main windows. Using a remote hook the dll is injected into each process.
The hook procedure gets the COM-Thread-ID and sends it to the utility. The utility corresponds the COM-Thread-Id with COM-Thread-ID from registry if they match it injects the DLL to this thread.
It creates the windows to communication and container. All the requests are transmitted through a proxy object which uses windows messages to communicate between threads.
The shared data is used to transmit the amount of data between processes without copying.

Thank you all for attention!

Best Regards,

The trick.
