Attribute VB_Name = "dx7subs"

'----DIRECTX----
Public DirectX As DirectX7                         'Main DirectX object
'---------------

'----DIRECTDRAW----
Public DirectDraw As DirectDraw7                   'Main DirectDraw object
Public Primary As DirectDrawSurface7               'Primary surface
Public BackBuffer As DirectDrawSurface7            'BackBuffer surface
Public SurfaceDesc As DDSURFACEDESC2               'Surface description
Public SurfaceDescZero As DDSURFACEDESC2           'Surface description (empty)
'------------------

'----DIRECT3D----
Public Direct3D As Direct3D7                       'Main Direct3D object
Public Device As Direct3DDevice7                   'Device
Public TempVertices(3) As D3DTLVERTEX              'TL Vertices
Public TempVector As D3DVECTOR                     'Vector
'----------------

Public BlendVertices As Single

Public SCREEN_WIDTH As Integer
Public SCREEN_HEIGHT As Integer

       
