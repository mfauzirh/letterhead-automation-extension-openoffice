Sub main

Dim sheet as Object
Dim tittle as String
Dim address as String
Dim contact as String
Dim path as String
Dim imageName as String

' only run extension in spreadsheet
if mfCheckComponent(thisComponent) then
	' get active sheet
	sheet = ThisComponent.getcurrentcontroller.activesheet
	
	' input data for header
	mfInputData(tittle, address, contact, path, imageName)
	
	' create header in sheet based inputed data
	mfCreateHeader(sheet, tittle, address, contact)
	
	' draw logo
	mfDrawLogo(path, imageName)
else
	msgbox "Hanya bisa berjalan di spreadsheet"
end if

End Sub

function mfCheckComponent(oDoc) as boolean
   if oDoc.supportsService("com.sun.star.sheet.SpreadsheetDocument") then
      mfCheckComponent = True
   else
      mfCheckComponent = False
   end if
End function

function mfInputData(ByRef refTittle, ByRef refAddress, ByRef refContact, ByRef refPath, ByRef refImageName)
	Dim inputText As String
	
	' input requirement data
	refTittle = InputBox ("Masukan Judul: ")
	refAddress = InputBox ("Masukan Alamat: ")
	refContact = InputBox ("Masukan Kontak (Email/NoTelepon): ")
	refPath = InputBox ("Masukan Path Gambar (C:\User\Downloads) (Kosongkan untuk default di C:\): ")
	refImageName = InputBox ("Masukan Nama Gambar (dengan ektensi gambar): ")
End function

function mfCreateHeader(ByRef refSheet, ByVal dataTittle, ByVal dataAddress, ByVal dataContact)
	' insert data into cell
	tittleText = refSheet.getCellRangeByName("B1")
	tittleText.String = dataTittle
	
	addressText = refSheet.getCellRangeByName("B2")
	addressText.String = dataAddress
	
	contactText = refSheet.getCellRangeByName("B3")
	contactText.String = dataContact
	
	' merge cell
    logo = refSheet.getCellrangeByname("A1:A3")
    logo.merge(True)

    title = refSheet.getCellrangeByname("B1:F1")
    title.merge(True)

    address = refSheet.getCellrangeByname("B2:F2")
    address.merge(True)

    contact = refSheet.getCellrangeByname("B3:F3")
    contact.merge(True)
End function

function mfDrawLogo(ByVal valPath, ByVal valImageName)
Dim Folder as String

if Len(path) <= 0 then
	Folder = "C:\"
else
	Folder = valPath
end if

imagen = valImageName
ImagenURL = convertToURL(Folder & imagen)
oImagen_obj = ThisComponent.createInstance("com.sun.star.drawing.GraphicObjectShape")

oImagen_obj.GraphicURL = ImagenURL
oSize = oImagen_obj.Size
oSize.Height = 1500
oSize.Width = 1500
oImagen_obj.Size = oSize
oPos = oImagen_obj.Position
oPos.X = 400
oPos.Y = 0
oImagen_obj.Position = oPos

oDP = ThisComponent.DrawPages.getByIndex(0)
oDP.add(oImagen_obj)

End function
