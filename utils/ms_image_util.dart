
import 'dart:io';
import 'dart:typed_data';
import 'package:image/image.dart' as img;


class MsImageUtil {

  MsImageUtil._({required this.image, this.wantedWidth, this.wantedHeight, this.aspectRatio = true});

  static Future<MsImageUtil> createInstance({required File image, int? wantedWidth, int? wantedHeight, bool aspectRatio = true}) async {
    
    // create ms util
    MsImageUtil msImageUtil = MsImageUtil._(
      image: image,
      wantedWidth: wantedWidth,
      wantedHeight: wantedHeight,
      aspectRatio: aspectRatio);
    
    // Set original width and height
    await msImageUtil.getOriginalImageDimensions();
    return msImageUtil;

  }


  File image;

  int? wantedWidth;
  int? wantedHeight;

  int? _originalWidth;
  int? _originalHeight;



  bool aspectRatio;


  // Setters and Getters

  int get width {
    Map<String,int> dimensions = msFindImageDimensions();
    return dimensions["width"]!;
  }

  int get height {
    Map<String,int> dimensions = msFindImageDimensions();
    return dimensions["height"]!;
  }



  // Helper Functions
  Map<String, int> msFindImageDimensions() {

    if(!aspectRatio){
      return {
        "width" : wantedWidth != null ? wantedWidth! : _originalWidth!,
        "height" : wantedHeight != null ? wantedHeight! : _originalHeight!,
      };
    }

    
    double widthScale = wantedWidth != null ? wantedWidth! / (_originalWidth!) : 1;
    double heightScale = wantedHeight != null ? wantedHeight! / (_originalHeight!) : 1;

    double scaleFactor = widthScale < heightScale ? widthScale : heightScale;

    // Calculate new dimensions
    int newWidth = (_originalWidth! * scaleFactor).toInt();
    int newHeight = ( _originalHeight! * scaleFactor).toInt();

    return {
      "width" : newWidth,
      "height" : newHeight,
    };

  }


  Future<Map<String, int>> getOriginalImageDimensions() async {
    if(_originalWidth != null && _originalHeight != null){
      return {
        "width" : _originalWidth!,
        "height" : _originalHeight!
      };
    }

    // Find original width and height from the file
    try {
      // Read the image file into bytes
      List<int> imageBytes = await image.readAsBytes();
      
      // Decode the image
      img.Image decodedImage = img.decodeImage(Uint8List.fromList(imageBytes))!;

      // Get image dimensions
      int width = decodedImage.width;
      int height = decodedImage.height;
      // print('Image width: $width, height: $height');

      _originalWidth = width;
      _originalHeight = height;

      return {
        "width" : width,
        "height" : height
      };

    } catch (e) {
      throw Exception("Error reading image $e");
    }

  }

  static int cmToEmu(double cm) {
    return (cm * 360000).round();
  }
  


}