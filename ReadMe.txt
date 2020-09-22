IMPORTANT--  Uses layered windows for transparency so is only windows 2000 and XP compatible. If you want to try it on other platforms comment out the lines specified in form_load before running.  Also you must have a recent flash ocx 6 or 7 to run.

  Originally made for my kids around christmas (explains the snowman).  Its a very simple animation but you can easily change it to your own.  It wanders around your desktop and can interact with the windows it finds.  It can be dragged, set on top of other windows, and even can detect when it bumps into another DesktopMate. (though doesn't do anything special right now when it finds one just a kindof annoying message box) Thanks to a variety of places (which I don't remember to give credit to)  for the window on top and drag code.

If you want to change the animation just change the movie property in shockwaveflash1 to your movie and change the constants in the general section to the frames each section of your animation begins at. --example LEFT_FRAME specifies which frame to play when the animation walks to the left.  you need at a minimum animations of walking, falling, when it lands, and, climbing.

  In the flash animation, put this actionscript at the end of each section:
    fscommand("Done");
it's to alert the vb part of this that the animations done if you use setWait.  Use that if you want an animation to complete fully, otherwise it will only play until it lands etc. and will play that animation.
  Then after the fscommand you can put whatever you want, like gotoAndPlay(1); or Stop();
  
  Have fun with it,
                  Mike


