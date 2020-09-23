'------------------------------
' Splitter Control v1.1.9
'
' by Tim Humphrey
'------------------------------

Splitter Control is a Windows Explorer style splitter.  To use it, add controls
to it the same way you do for a frame, and set the Child1 and/or Child2
properties to the names of the added controls.

Features
--------

- Live and non-live updating while child controls are resized
- Customize appearance of splitter bar while not live updating
- Escape can be pressed while moving splitter bar to cancel the move
- Specify a maximum size for a child control
- Specify a minimum size for either child control
- Horizontal or vertical orientation
- Position splitter bar by percentage or absolute position

Notes
-----

- Documentation of the Splitter.ocx is provided courtesy of the
ActiveX Documenter, available free from http://www.vbaccelerator.com

- Splitter.vbp is the project that actually compiles the .ocx file.

- Test.vbp is a project I used to test the control, it also doubles as the
demo project.

- Both the Spitter and Test projects use the same user control and both require
the public property to be set to different values; VB will inform you of this
if the value should be changed.  The public property should be true in the
Splitter project and false in the Test project.

Version History
---------------

1.1.9 - 10/15/2001
- Added AllowResize property per a user request.  This allows you to hide the
  splitter bar in that the cursor won't change to an arrow when the user moves
  their mouse over it.  Note, the splitter bar can still be programmatically
  moved when hidden in this way.
- Removed the restriction on the minimum size of the splitter bar.  The
  SplitterSize property can now be as small as 0.

1.1.8 - 4/19/2001

* Started all enumerations at 0, breaks compatibility with previous
  version on OrientationConstants enumeration.

  Sorry for doing this but I think it's better this way and the control is
  relatively young so this shouldn't hurt too many people.  If it does open
  your .frm file in Notepad and go to the line that contains "splitter.ocx".
  Change the GUID to {E4803A90-1A86-44D8-B884-F93EB01FC8E8} and save.
  (This number is also listed at the beginning of the documentation.)
  When you open the project VB should ask you to upgrade, say yes, it's
  changing the number that was within the # signs in the frm.  You'll also
  want to reset the orientation property once you open the project.

  I know breaking compatibility sucks and I don't plan on doing this again.

- Included documentation of the control in the file, Documentation.rtf
- Added BorderStyle, MaxSize, MaxSizeAppliesTo, SplitterPos, CurrSplitterPos,
  CurrRatioFromTop and Maintain properties
- Rewrote code, where necessary, to support new properties and to
  reduce complexity
- CurrSplitterPos and CurrRatioFromTop always report accurate readings;
  previously RatioFromTop assumed CurrRatioFromTop's functionality
  and could sometimes report inaccurate readings
- Out of necessity, gave invalid property values proper values
- Removed design-time splitter appearance to make control easier to grab
- Reduced default splitter size

1.0.0 - 3/14/2001

- Initial creation

--
Tim Humphrey
10/15/2001
zzhumphreyt@techie.com
