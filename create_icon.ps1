# Create an icon file for CA Manager
Add-Type -AssemblyName System.Drawing

# Create a 256x256 bitmap for high-quality icon
$bitmap = New-Object System.Drawing.Bitmap(256, 256)
$graphics = [System.Drawing.Graphics]::FromImage($bitmap)
$graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
$graphics.Clear([System.Drawing.Color]::White)

# Draw a shield shape (security/conditional access theme)
$shieldPath = New-Object System.Drawing.Drawing2D.GraphicsPath

# Shield points
$points = @(
    [System.Drawing.Point]::new(128, 20),   # Top center
    [System.Drawing.Point]::new(200, 60),  # Top right curve
    [System.Drawing.Point]::new(200, 180),  # Right side
    [System.Drawing.Point]::new(128, 236),  # Bottom center
    [System.Drawing.Point]::new(56, 180),  # Left side
    [System.Drawing.Point]::new(56, 60)    # Top left curve
)
$shieldPath.AddPolygon($points)

# Fill shield with blue gradient
$brush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
    [System.Drawing.Point]::new(0, 0),
    [System.Drawing.Point]::new(256, 256),
    [System.Drawing.Color]::FromArgb(0, 120, 215),
    [System.Drawing.Color]::FromArgb(0, 80, 160)
)
$graphics.FillPath($brush, $shieldPath)

# Draw shield border
$pen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(0, 100, 200), 8)
$graphics.DrawPath($pen, $shieldPath)

# Draw a checkmark inside (representing conditional access/approval)
$checkPen = New-Object System.Drawing.Pen([System.Drawing.Color]::White, 20)
$checkPen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
$checkPen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
$graphics.DrawLine($checkPen, 90, 130, 120, 160)
$graphics.DrawLine($checkPen, 120, 160, 170, 100)

# Draw "CA" text at bottom
$font = New-Object System.Drawing.Font("Arial", 48, [System.Drawing.FontStyle]::Bold)
$textBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
$graphics.DrawString("CA", $font, $textBrush, 85, 190)

# Clean up
$graphics.Dispose()

# Create icon with multiple sizes (16, 32, 48, 256)
$iconStream = New-Object System.IO.MemoryStream
$iconWriter = New-Object System.IO.BinaryWriter($iconStream)

# Write ICO header
$iconWriter.Write([UInt16]0)  # Reserved
$iconWriter.Write([UInt16]1)  # Type (1 = ICO)
$iconWriter.Write([UInt16]4)  # Number of images

# Function to resize bitmap
function Resize-Bitmap {
    param($source, $size)
    $resized = New-Object System.Drawing.Bitmap($size, $size)
    $g = [System.Drawing.Graphics]::FromImage($resized)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g.DrawImage($source, 0, 0, $size, $size)
    $g.Dispose()
    return $resized
}

# Write image entries and data for each size
$sizes = @(16, 32, 48, 256)
$imageData = @()

foreach ($size in $sizes) {
    $resized = Resize-Bitmap $bitmap $size
    
    # Convert to PNG in memory
    $pngStream = New-Object System.IO.MemoryStream
    $resized.Save($pngStream, [System.Drawing.Imaging.ImageFormat]::Png)
    $pngBytes = $pngStream.ToArray()
    $pngStream.Dispose()
    $resized.Dispose()
    
    # Write directory entry
    if ($size -eq 256) { 
        $iconWriter.Write([Byte]0)  # Width (0 = 256)
        $iconWriter.Write([Byte]0)  # Height (0 = 256)
    } else { 
        $iconWriter.Write([Byte]$size)  # Width
        $iconWriter.Write([Byte]$size)  # Height
    }
    $iconWriter.Write([Byte]0)  # Color palette
    $iconWriter.Write([Byte]0)  # Reserved
    $iconWriter.Write([UInt16]1)  # Color planes
    $iconWriter.Write([UInt16]32)  # Bits per pixel
    $iconWriter.Write([UInt32]$pngBytes.Length)  # Image data size
    $iconWriter.Write([UInt32]($iconStream.Length + 22 + ($sizes.Count * 16)))  # Offset (will fix later)
    
    $imageData += $pngBytes
}

# Calculate correct offsets
$baseOffset = 6 + ($sizes.Count * 16)  # Header + directory entries
$currentOffset = $baseOffset

# Rewrite directory entries with correct offsets
$iconStream.Position = 6
foreach ($i in 0..($sizes.Count - 1)) {
    $size = $sizes[$i]
    if ($size -eq 256) { 
        $iconWriter.Write([Byte]0)  # Width (0 = 256)
        $iconWriter.Write([Byte]0)  # Height (0 = 256)
    } else { 
        $iconWriter.Write([Byte]$size)  # Width
        $iconWriter.Write([Byte]$size)  # Height
    }
    $iconWriter.Write([Byte]0)
    $iconWriter.Write([Byte]0)
    $iconWriter.Write([UInt16]1)
    $iconWriter.Write([UInt16]32)
    $iconWriter.Write([UInt32]$imageData[$i].Length)
    $iconWriter.Write([UInt32]$currentOffset)
    $currentOffset += $imageData[$i].Length
}

# Write image data
foreach ($data in $imageData) {
    $iconWriter.Write($data)
}

$iconWriter.Flush()
$iconBytes = $iconStream.ToArray()
$iconWriter.Dispose()
$iconStream.Dispose()

# Save icon file
[System.IO.File]::WriteAllBytes("$PSScriptRoot\ca_manager.ico", $iconBytes)
$bitmap.Dispose()

Write-Host "Icon created successfully: ca_manager.ico" -ForegroundColor Green

