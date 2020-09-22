                        --------------------------------------
			Advanced Picturebox (cTransPicturebox)
			--------------------------------------
                                (c) OMIT, Inc.  2002


The perfect blend of a picture box and image control, the Advanced Picture Box is a lightweight control that makes applications such as games easier to produce.

- User-customizable transparency.  Either use the default transparency color or set your own to any other color
- Stretching like the image control
- Border style (on or off)
- Image tiling
- Transparency Synchronizing.  By default, the transparent color is the color at point (0,0)


PROPERTIES
==========

BorderStyle:		Sets/returns the style of the border around the control
			(0 - none, 1 - fixed [default])

PictureFile:		Sets/returns the picture used for the control. 
			If SyncTransparentColor is set to True, the color 
			at point (0,0) is automatically used for the 
			Transparent color

Style:			Sets/returns the picture display style
			(0 - none, 1 - repeat [default], 2 - stretch)

SyncTransparentColor:	Determines if the transparent color is automatically 
			chosen or can be manually set by the user.  When set to
			true, the "hotpoint" color (point (0,0) of the picture, 
			top-left corner) is used as the transparent color

TransparentColor:	The color that is made invisible on the picture.  This can
			be any color available to Windows.  TransparentColor is
			influenced by the PictureFile and SyncTransparentColor
			properties.

Common properties
=================

All other common properties function as a normal picturebox


NOTE
====

Unlike the picturebox, this control cannot be used as a container for other controls.
Although this control can accept BMP, GIF, JPG, TIF files, large picture files do consume memory.


USAGE
=====

ActiveX (OCX) version:	Add this control by selecting it in the Components dialog
			(Project>Components...)

UserControl version:	Add this control to your project by selecting the 					cTransPicturebox.ctl UserControl file

Control Prefix:		tpb (ie. tpbMyPic.name)

