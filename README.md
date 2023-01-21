<div align="center">
<img width="418" height="463" src="https://dl.unicontsoft.com/upload/pix/ss_qr_code5.png">

## VbQRCodegen
QR Code generator library for VB6/VBA
</div>

### Description

A single file QR Code generator based on https://www.nayuki.io/page/qr-code-generator-library

### Usage

Just add `mdQRCodegen.bas` to your project and call `QRCodegenBarcode` function to retrieve an picture from a text or a byte-array like this:

```
    Set Image1.Picture = QRCodegenBarcode("Sample text")
```
Note that you can stretch/zoom the returned picture to any size without loss of quality because the picture is using vectors to draw the QR Code.

### MS Access Support

For compatibility with image controls on forms/reports you can use `QRCodegenConvertToData` function like this:
```
    Image0.PictureData = QRCodegenConvertToData(QRCodegenBarcode("Sample text"), 500, 500)
```
Note that this produces bitmap picture of the QR Code so might need to tweak output size parameters.