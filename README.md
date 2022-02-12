# VbVst
## The VB6 VST Framework
 
 Hello everyone!
 
 This framework allows you to create VST 2.X plugins (so far only the effects) in VB6. This doesn't support all the 2.4 features yet.
 To create a VST plugin you need to perform the several steps:
 
 1. Install **VBCDeclFix Add-in** (https://github.com/thetrik/VBCDeclFix);
 2. Create an **ActiveX DLL** project;
 3. Specify the single-threading model in the **Project->Properties->General** tab;
 4. Add **modVBVst2X.bas** module to project (in **vbvst2x** directory);
 5. Add **vbvst2x.tlb** to the **Project->References** (in the **vbvst2x\TypeLib** directory);
 6. Add a public creatable class. Add global constant **VST_PLUGIN_CLASS_NAME** with its name;
 7. You should impement **IVBVstEffect** interface in this class;
 8. Export **VSTPluginMain** function with names **VSTPluginMain** and **main** using the following link switches: **-EXPORT:VSTPluginMain -EXPORT:main=VSTPluginMain** http://bbs.vbstreets.ru/viewtopic.php?f=9&t=43618 (VBCompiler/LinkSwitches);
 9. Add the effect logic;
 10. Compile project to native code with all the optimizations enabled.
 
 To debug your VST plugin you can use the **VbDebugVst** project in this repository. This is the very small-basic DAW utility which allows you to debug your VST plugins inside IDE. You only need to specify **ProgID** of your class and connect to IDE. This utility can pass a sound and some MIDI and automation events through your plugin, save/load state. See the description for more information inside **VbDebugVst** directory.
 
 You can use the template from **template** directory which conatains the project you can start with. You can drop this project to the **VB98\Template\Projects** directory (don't forget to update **modVBVst2X.bas** location) and use the new project type.
 
 This repository contains 2 complete VST plugins written using this framework (**VbTrickGlitch** and **VbTrickCrusher** directories). You can read description in the corresponding directories.
 
 I've tested the produced plugins in several DAWs and everything worked fine but the bugs are possible because this is poorly tested.
 
 ## How does it work?
 
 The module **modVBVst2X** contains the special marshaler and the project/runtime initializer. So because of VB6 object instances live in single apartment you need to initialize apartment for your object and route all the calls from different threads to object's STA. The module performs this task. It creates the new STA for your objects and initializes the project context. After that, you can safe work with the objects. The marshaler is based on windows messages. It just sends the messages from the arbitrary threads to the object's STA. To avoid redistribute tlb i decided to use manual marshaling but when i started developing this project i used typelib marshaller, just FYI. The initial purpose of this project was **VbTrickGlitch** plugin which i currently use then i just designed it as a framework, added the simple project (**VbTrickCrusher**), the template, created the debugger. This isn't the first VST plugins written in VB6 by me. I did it before (https://youtu.be/AKDSd5J7pgY) but the main advantage here is **VBCDeclFix** which allows to create **cdecl** callbacks functions which are required for VST. 
 
 Thank you all for attention!
 
 Best Regards,

 The trick.
