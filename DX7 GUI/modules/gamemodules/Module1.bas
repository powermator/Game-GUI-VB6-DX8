Attribute VB_Name = "dxobjects"
'----------------------------------------------------dx object

Public dx                  As New DirectX7


'---------------------------------------------------ddraw
Public Dd                  As DirectDraw4
Public Primary             As DirectDrawSurface4
Public Back_Buffer         As DirectDrawSurface4
Public Ddsd                As DDSURFACEDESC2


'----------------------------------------------------d3drm3
Public D3d                 As Direct3DRM3
Public Device              As Direct3DRMDevice3
Public ViewPort            As Direct3DRMViewport2

Public WorldFrame          As Direct3DRMFrame3
Public CameraFrame         As Direct3DRMFrame3

Public Shadow              As Direct3DRMLight
Public Light               As Direct3DRMLight

Public LightFrame          As Direct3DRMFrame3


Public MeshFrame(10)       As Direct3DRMFrame3
Public Mesh(10)            As Direct3DRMMeshBuilder3

Public MeshFramex     As Direct3DRMFrame3
Public Meshx          As Direct3DRMMeshBuilder3


Public dxv(10)            As D3DLVERTEX
Public dxf(10)            As Direct3DRMFace2

Public tex                As Direct3DRMTexture3
Public tex2               As Direct3DRMTexture3
Public tex3               As Direct3DRMTexture3





Public Trans_Value        As Single

