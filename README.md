<div align="center">

## The CORRECT Universal Exit Code


</div>

### Description

This article discusses the PROPER ways to end a program.. I can't believe I'm actually posting this, but after seeing another posting on here and responses where they refer to properly ending a program as "The Advanced Way".... Well.. Here we go..
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[The Chameleon](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/the-chameleon.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/the-chameleon-the-correct-universal-exit-code__1-65734/archive/master.zip)





### Source Code

Perhaps you're new to programming and are having problems shutting down your application properly? Every day I see many people use, and recommend using, the simple command <b>End</b><br><br>
<i>You should NOT use End</i> To quote from a comment on the VB Wiki:<br><br>
"End is the equivalant of stopping your car by driving it into a brick wall (stolen from a DevX newsgroup years ago :o). It does NO clean up at all, it stops the process dead. This is especially bad if you are doing any form of database or hardware access."<br><br>
Adding in my own comments to this, End is one of several commands that cause Visual Basic to be labeled as being horrible and inefficient in the programming community.. Commands that result in a program running improperly, using great resources, and (in this case) leaving memory leaks..<br><br>
An application will shut down once <i>All objects have been unloaded</i> and <i>All code has stopped running</i>.. Unloading your objects (usually just forms) is simple.. You just Unload frmName.. You can even put it all in the QueryUnload event of your main form if you want all other forms to unload with it..<br><br>
If you're in the VB IDE and your program isn't closing, check to see if a form's the culprit first.. If you run <b>For Each Form In Forms: Print Form.Name: Next</b> in your Intermediate window, you'll see a listing of all of your currently loaded forms.<br><br>
You can also loop through all of your forms and unload all of them by:<br>
<b>
Dim Form As Form<br>
For Each Form In Forms<br>
&nbsp; Unload Form<br>
Next<br></b><br>
Aside from forms, you may be loading other objects manually, User Interfaces or otherwise.. Be sure to unload all of these..<br><br>
And, lastly, you may have a tight loop somewhere in your code.. You shouldn't have one without knowing it's there, so if you don't know what I'm talking about, don't worry..<br><br>
In a tight loop, you should ALWAYS make a boolean that can stop it.. For example, a boolean named ProgramClosing.. Set it to true in your QueryUnload event, and also check in your loop and if FormClosing then exit the loop and stop running..<br><br>
Hopefully this will help clear up some things for someone...

