Imports TIS.Imaging

Public Class Form1
    Dim VideoHasStarted As Boolean
    Dim StartToggle As Integer = 0
    Dim ZoomOutValue As Decimal = 1.2
    Dim TriggerToggle As Integer = 0
    Dim ImgSaveToggle As Integer = 0
    Dim BitToggle As Integer = 0
    Dim TiffToggle As Integer = 0
    Dim impath_hi As String
    Dim impath_lo As String
    Dim imFullpath_hi As String = "C:\"
    Dim imFullpath_lo As String = "C:\"
    Dim IncrementLow As Decimal = 0
    Dim ImageCount As Integer = 0
    Dim SequenceCount As Integer = 0
    Dim FrameCount As Integer = 0
    Dim ExposureProperty As VCDRangeProperty
    Dim ExposureAuto As VCDSwitchProperty
    Dim ExposureMax As VCDRangeProperty
    Dim ExposureMaxAuto As VCDSwitchProperty
    Dim ImageOdd As Integer = 0
    Dim ImageEven As Integer = 0
    Dim ImageSaveCount As Integer = 0

    Private _imageSink As TIS.Imaging.FrameSnapSink
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        On Error GoTo err_Form_Load
        If Not IcImagingControl1.DeviceValid Then
            IcImagingControl1.ShowDeviceSettingsDialog()

            If Not IcImagingControl1.DeviceValid Then
                MsgBox("No device was selected.")

                Exit Sub
            End If
        End If
        IcImagingControl1.LiveDisplay = True
        IcImagingControl1.LiveCaptureLastImage = False
        IcImagingControl1.LiveCaptureContinuous = True


        IcImagingControl1.Width = IcImagingControl1.ImageWidth
        IcImagingControl1.Height = IcImagingControl1.ImageHeight
        'IcImagingControl1.MemoryCurrentGrabberColorformat = ICImagingControlColorformats.ICRGB64
        IcImagingControl1.MemoryCurrentGrabberColorformat = ICImagingControlColorformats.ICRGB32

        'use this for tiff images
        '_imageSink = New TIS.Imaging.FrameSnapSink(TIS.Imaging.MediaSubtypes.RGB64)
        'IcImagingControl1.Sink = _imageSink
        'IcImagingControl1.ImageAvailableExecutionMode = 1

        Exit Sub

err_Form_Load:
        MsgBox(Err.Description)
    End Sub
    Private Sub ShowBuffer(buffer As IFrameQueueBuffer)
        Try
            IcImagingControl1.DisplayImageBuffer(buffer)
        Catch
            MessageBox.Show("snap image failed, timeout occurred.")
        End Try
    End Sub






    Private Sub IcImagingControl1_Load(sender As Object, e As EventArgs) Handles IcImagingControl1.Load

    End Sub

    Private Sub Start_Click(sender As Object, e As EventArgs) Handles Start.Click
        If StartToggle = 0 Then
            Start.BackColor = Color.Red
            StartToggle = 1
            IcImagingControl1.LiveStart()
            VideoHasStarted = True
        Else
            Start.BackColor = Color.White
            StartToggle = 0
            IcImagingControl1.LiveStop()
            VideoHasStarted = False
        End If
    End Sub

    Private Sub zoomout_click(sender As Object, e As EventArgs) Handles ZoomOut.Click
        IcImagingControl1.LiveDisplayDefault = 0
        IcImagingControl1.LiveDisplayZoomFactor = IcImagingControl1.LiveDisplayZoomFactor / ZoomOutValue
    End Sub

    Private Sub zoomin_click(sender As Object, e As EventArgs) Handles ZoomIn.Click
        IcImagingControl1.LiveDisplayDefault = 0
        IcImagingControl1.LiveDisplayZoomFactor = IcImagingControl1.LiveDisplayZoomFactor * ZoomOutValue
    End Sub

    Private Sub settings_click(sender As Object, e As EventArgs) Handles Settings.Click
        IcImagingControl1.ShowPropertyDialog()
    End Sub

    Private Sub trigger_click(sender As Object, e As EventArgs) Handles Trigger.Click
        If TriggerToggle = 0 Then
            Trigger.BackColor = Color.Red
            TriggerToggle = 1
            IcImagingControl1.DeviceTrigger = True
        Else
            Trigger.BackColor = Color.White
            TriggerToggle = 0
            IcImagingControl1.DeviceTrigger = False
        End If
    End Sub

    Private Sub imgsave_click(sender As Object, e As EventArgs) Handles ImgSave.Click
        If ImgSaveToggle = 0 Then
            ImgSave.BackColor = Color.Red
            ImgSaveToggle = 1
        Else
            ImgSave.BackColor = Color.White
            ImgSaveToggle = 0
        End If

    End Sub

    Private Sub path1_click(sender As Object, e As EventArgs) Handles Path1.Click
        If impath_hi = "" Then
            Dim dialog As New FolderBrowserDialog()
            dialog.RootFolder = Environment.SpecialFolder.Desktop
            dialog.SelectedPath = "c:\"
            dialog.Description = "select application configuration files path"
            If dialog.ShowDialog() = Windows.Forms.DialogResult.OK Then
                impath_hi = dialog.SelectedPath
            End If
            'my.computer.filesystem.writealltext(impath & "apppath.txt", impath, false)
            Debug.Print(impath_hi)
        End If
    End Sub

    Private Sub path2_click(sender As Object, e As EventArgs) Handles Path2.Click
        If impath_lo = "" Then
            Dim dialog As New FolderBrowserDialog()
            dialog.RootFolder = Environment.SpecialFolder.Desktop
            dialog.SelectedPath = "c:\"
            dialog.Description = "select application configeration files path"
            If dialog.ShowDialog() = Windows.Forms.DialogResult.OK Then
                impath_lo = dialog.SelectedPath
            End If
            'my.computer.filesystem.writealltext(impath & "apppath.txt", impath, false)
            Debug.Print(impath_lo)
        End If
    End Sub

    Private Sub inclow_click(sender As Object, e As EventArgs) Handles IncLow.Click
        MessageBox.Show("Exposure increment=" + Str(IncrementLow))
    End Sub


    Private Sub icimagingcontrol1_imageavailable(sender As Object, e As ICImagingControl.ImageAvailableEventArgs) Handles IcImagingControl1.ImageAvailable

        On Error GoTo err_imageavailable_handler

        Dim CurrentBuffer As ImageBuffer
        Dim DisplayBuffer As ImageBuffer
        Dim bufferNumber As Integer = IcImagingControl1.ImageBuffers.Count
        CurrentBuffer = IcImagingControl1.ImageActiveBuffer

        ExposureProperty = IcImagingControl1.VCDPropertyItems.FindInterface(VCDGUIDs.VCDID_Exposure, VCDGUIDs.VCDElement_Value, VCDGUIDs.VCDInterface_Range)
        ExposureAuto = IcImagingControl1.VCDPropertyItems.FindInterface(VCDGUIDs.VCDID_Exposure, VCDGUIDs.VCDElement_Auto, VCDGUIDs.VCDInterface_Switch)
        ExposureMax = IcImagingControl1.VCDPropertyItems.FindInterface(VCDGUIDs.VCDID_Exposure, VCDGUIDs.VCDElement_AutoMaxValue)
        ExposureMaxAuto = IcImagingControl1.VCDPropertyItems.FindInterface(VCDGUIDs.VCDID_Exposure, VCDGUIDs.VCDElement_AutoMaxValueAuto, VCDGUIDs.VCDInterface_Switch)
        ExposureMaxAuto.Switch = False
        ExposureMax.Value = 1973

        If ImgSaveToggle = 1 Then
            FrameCount = FrameCount + 1
            ImageCount = ImageCount + 1
            SequenceCount = SequenceCount + 1
            Debug.Print(ImageCount)
            If SequenceCount = 5 Then
                SequenceCount = 1
            End If
            If ImageCount > 1 Then
                If SequenceCount = 1 Then
                    'Threading.Thread.Sleep(50)
                    ExposureAuto.Switch = False
                    ExposureProperty.Value = ExposureProperty.Value - IncrementLow
                    ImageOdd = ImageOdd + 1
                    ImageSaveCount = ImageSaveCount + 1
                    impath_lo = impath_lo.Trim(vbNullChar)
                    If TiffToggle = 1 Then
                        'imFullpath_hi = impath_hi + "\" + Str(ImageOdd) + ".tiff"
                        imFullpath_hi = impath_hi + "\" + Str(ImageSaveCount) + ".tiff"
                    Else
                        'imFullpath_hi = impath_hi + "\" + Str(ImageOdd) + ".jpg"
                        imFullpath_hi = impath_hi + "\" + Str(ImageSaveCount) + ".jpg"
                    End If
                    imFullpath_hi = imFullpath_hi.Replace(" ", "")
                    If TiffToggle = 1 Then
                        CurrentBuffer.SaveAsTiff(imFullpath_hi)
                    Else
                        CurrentBuffer.SaveAsJpeg(imFullpath_hi, 100)
                    End If
                End If
                If SequenceCount = 3 Then
                    'Threading.Thread.Sleep(50)
                    ExposureProperty.Value = ExposureProperty.Value + IncrementLow
                    ExposureAuto.Switch = True
                    ImageEven = ImageEven + 1
                    ImageSaveCount = ImageSaveCount + 1
                    impath_hi = impath_hi.Trim(vbNullChar)
                    If TiffToggle = 1 Then
                        'imFullpath_lo = impath_lo + "\" + Str(ImageEven) + ".tiff"
                        imFullpath_lo = impath_lo + "\" + Str(ImageSaveCount) + ".tiff"
                    Else
                        'imFullpath_lo = impath_lo + "\" + Str(ImageEven) + ".jpg"
                        imFullpath_lo = impath_lo + "\" + Str(ImageSaveCount) + ".jpg"
                    End If
                    imFullpath_lo = imFullpath_lo.Replace(" ", "")
                    If TiffToggle = 1 Then
                        CurrentBuffer.SaveAsTiff(imFullpath_lo)
                    Else
                        CurrentBuffer.SaveAsJpeg(imFullpath_lo, 100)
                    End If
                End If
                'Debug.Print("image count=" + CStr(ImageCount))
                'Debug.Print("imageodd=" + Str(ImageOdd))
                'Debug.Print("imageeven=" + Str(ImageEven))
                'Debug.Print(imFullpath_hi)
                'Debug.Print(imFullpath_lo)
            End If
        End If
err_imageavailable_handler:
        Debug.Print(Err.Description)
    End Sub

    Private Sub SaveConf_Click(sender As Object, e As EventArgs) Handles SaveConf.Click
        IcImagingControl1.SaveDeviceStateToFile("device_state.txt")
    End Sub

    Private Sub LoadConf_Click(sender As Object, e As EventArgs) Handles LoadConf.Click
        IcImagingControl1.LoadDeviceStateFromFile("device_state.txt", True)
    End Sub

    Private Sub Bit64_Click(sender As Object, e As EventArgs) Handles Bit64.Click
        If BitToggle = 0 Then
            Bit64.BackColor = Color.Red
            BitToggle = 1
            IcImagingControl1.MemoryCurrentGrabberColorformat = ICImagingControlColorformats.ICRGB64
        Else
            Bit64.BackColor = Color.White
            BitToggle = 0
            IcImagingControl1.MemoryCurrentGrabberColorformat = ICImagingControlColorformats.ICRGB24
        End If
    End Sub

    Private Sub SaveTiff_Click(sender As Object, e As EventArgs) Handles SaveTiff.Click
        If TiffToggle = 0 Then
            SaveTiff.BackColor = Color.Red
            TiffToggle = 1
        Else
            SaveTiff.BackColor = Color.White
            TiffToggle = 0
        End If
    End Sub
    Private Sub NumericUpDown1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NumericUpDown1.ValueChanged
        'NumericUpDown valu will show in TextBox
        IncrementLow = NumericUpDown1.Value / 10
    End Sub
    Private Sub NumericUpDown2_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NumericUpDown2.ValueChanged
        'NumericUpDown valu will show in TextBox
        ImageSaveCount = NumericUpDown2.Value
    End Sub
    Private Sub showimgnum_click(sender As Object, e As EventArgs) Handles ShowImgNum.Click
        MessageBox.Show("Img start number =" + Str(ImageSaveCount))
    End Sub
End Class
