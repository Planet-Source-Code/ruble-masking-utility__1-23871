Attribute VB_Name = "Module1"
'******************************************************************
'* Made by Ruble
'* ruble_19@yahoo.com
'**********************
'*
'* You may use this code any way you like.
'*
'* Description:
'* This program takes an image and makes a mask image of it
'* It uses black as the masking color
'* Everything that is black in the original will be masked
'*
'* Usage:
'* 1) Double Click on the original picture box to load the picture
'* 2) Click the "MASK" Button
'* 3) Click the "TEST" Button to test if the mask works
'* 4) Click the "STOP" Button to stop the test
'* 5) Click the "SAVE" Button to save the mask image
'* 6) Repeat to make another mask
'*
'* Warning!!!
'* For very big pictures the form will not resize, however it will still work
'******************************************************************



'BitBlt in order to test the mask
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As _
        Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal _
        hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
