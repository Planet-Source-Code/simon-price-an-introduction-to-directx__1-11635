VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "DirectX Tutorial #1 - DirectDraw Intro - by Simon Price"
   ClientHeight    =   4788
   ClientLeft      =   48
   ClientTop       =   324
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4788
   ScaleWidth      =   6000
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   4572
      Left            =   120
      ScaleHeight     =   4524
      ScaleWidth      =   5724
      TabIndex        =   0
      Top             =   120
      Width           =   5772
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''
'
'         DX INTRO - BY SIMON PRICE
'
'     AN TASTE OF DIRECTX FOR BEGINNERS
'
''''''''''''''''''''''''''''''''''''''''''''''

' This is the main DirectX object, it contains
' functions that will allow you to create other
' main objects such as DirectDraw, Direct3D
' DirectInput and DirectSound.

Dim DX As New DirectX7

' This is the main DirectDraw object. It has all
' the graphics functions you need. It is created
' by DirectX.

Dim DDRAW As DirectDraw7

' This represents the primary surface. The primary
' surface is what you see on the screen, so if
' you draw on this, it will appear on the screen.

Dim Primary As DirectDrawSurface7

' This describes a surface so DirectDraw knows it's
' properties

Dim SurfDesc As DDSURFACEDESC2

' This represents the memory where we will store
' a bitmap. We can load a .bmp file and it will
' be put onto this surface.

Dim picBMP As DirectDrawSurface7

' This is a DirectDrawClipper. It allows you to
' draw onto a surface without worrying about
' where the edge of the screen is. If you try to
' draw off the screen, the clipper will stop any
' errors occuring by "clipping" the image to fit
' on the surface.

Dim Clipper As DirectDrawClipper


' in the form load event, we will start DirectDraw
' and load our picture from the file "test.bmp"

Private Sub Form_Load()
' if an error occurs, we will exit the program.
' in future tutorials, we will learn more about
' what errors can occur and how we can handle
' them without ending the program
On Error GoTo StartUpFailed

' set the scalemode to pixels
ScaleMode = 3

' ask DirectX to create DirectDraw
Set DDRAW = DX.DirectDrawCreate("")

' set the cooperative level - we only need a normal
' cooperative level, in future tutorials we will
' learn about exclusive mode which gives us more
' power than normal mode
DDRAW.SetCooperativeLevel Me.hWnd, DDSCL_NORMAL

' set the surface description flags to show it is
' a primary surface
SurfDesc.lFlags = DDSD_CAPS
SurfDesc.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE

' create the primary surface, telling DirectDraw
' the surface description
Set Primary = DDRAW.CreateSurface(SurfDesc)

' now we make the description for the other surface.
' it is an offscreen surface, meaning that we can't
' see it on the screen
SurfDesc.lFlags = DDSD_CAPS
SurfDesc.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN

' now we load the picture file to create our second
' surface. Remember, we can't see this, it's just
' in memory
Set picBMP = DDRAW.CreateSurfaceFromFile(App.Path & "\test.bmp", SurfDesc)

' now we create the clipper. this stops us from drawing
' outside the picture and making a mess all over the
' screen
Set Clipper = DDRAW.CreateClipper(0)

' this sets the clipper to allow drawing only in the
' picturebox
Clipper.SetHWnd Picture1.hWnd

' this tells the primary surface to use the clipper
' we have just made
Primary.SetClipper Clipper

' we're done, we don't want to go to the next bit,
' so exit the sub
Exit Sub

' this is where we go if there's an error
StartUpFailed:

' display an error message
MsgBox "ERROR : StartUp failed!", vbCritical, "ERROR!"

' end program
Unload Me
End Sub

Private Sub Form_Resize()
' this tells the picturebox to change size with
' the form
Picture1.Width = Me.ScaleWidth - 20
Picture1.Height = Me.ScaleHeight - 20

' repaint the picture to fit new size
Picture1.Refresh

End Sub

' the picturebox's paint event occurs when it is shown
' and so we need to paint something intersting in it.
' here we paint out picture stored in the picBMP
' surface

Private Sub Picture1_Paint()
' we need to store the rectangle sizes of the
' picture and the picturebox, so we can tell
' DirectDraw to stretch it to fit in
Dim destRect As RECT
Dim srcRect As RECT

' we can ask DirectX to get the size of the
' picturebox for us
DX.GetWindowRect Picture1.hWnd, destRect

' we can set the size of the source rectangle
' to the size of the picture surface by getting
' and using it's description
picBMP.GetSurfaceDesc SurfDesc
srcRect.Right = SurfDesc.lWidth
srcRect.Bottom = SurfDesc.lHeight

' now we use the blt function, this copies one
' picture to another, and stretches it to fit
' if necessary. here we tell it to blt from our
' picture surface (which we can't see) to the
' picturebox on the primary surface (which we can
' see), using the rectangle sizes to fit the
' picture. we also tell it to wait (DDBLT_WAIT)
' because a blt cannot always be performed
' immediately, so it's safer to let it wait until
' it's OK to blt.
Primary.Blt destRect, picBMP, srcRect, DDBLT_WAIT

End Sub
