import 'dart:io';
import 'package:archive/archive.dart';
import 'package:flutter/services.dart';
import 'package:path/path.dart' as path;
import 'package:path_provider/path_provider.dart';
import 'package:tgb_danismanlik_ui/const/ms_word_constant.dart';
import 'package:tgb_danismanlik_ui/utils/ms_image_util.dart';
import 'package:xml/xml.dart' as xml;

/// A utility class for manipulating DOCX files, providing functionality for text replacement,
/// image insertion, and document manipulation.
class DocxUtil {


  /// Creates a new instance of DocxUtil.
  /// 
  /// [filePath] The path to the DOCX file to be manipulated
  /// [savePath] The path where the modified DOCX file will be saved
  /// [isAsset] Whether the file is an asset (true) or a local file (false)
  DocxUtil(
      {required this.filePath, required this.savePath, required this.isAsset}) {
    if (!filePath.toLowerCase().endsWith('.docx')) {
      throw Exception("The file at $filePath is not a .docx file.");
    }

    isExtracted = extractDocx(filePath, isAsset);
  }

  /// Cleans up temporary files and resources used by the DocxUtil instance.
  void dispose() {
    final tempDirFolder = Directory(tmpPath!);

    // Clean up by deleting the temporary directory and its contents
    tempDirFolder.deleteSync(recursive: true);
    xmlContentMap.clear();
  }

/*
  !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
  Variables
  !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
*/

  static const tmpDirName = "TGB_DANISMANLIK_UI_TMP";

  String filePath;
  String savePath;
  bool isAsset;
  String? tmpPath;
  late Future<bool> isExtracted;

  Map<String, int> currentRIdNumberMap = {};
  Map<String, String> xmlContentMap = {};

/*
  !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
  Main Functions
  !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
*/

  /// Replaces text placeholders in the document with provided values.
  /// 
  /// [textReplacements] A map of placeholder keys to their replacement values
  Future<void> replaceText(
      {required Map<String, dynamic> textReplacements}) async {
    // Wait for the file to be extracted
    await isExtracted;

    // Read XML Content
    String documentFilePath = generateXMLPathDocument();

    // Replace placeholders
    _replacePlaceholders(
        documentPath: documentFilePath, replacements: textReplacements);
  }

  /// Replaces entire paragraphs containing placeholders with provided values. It should be used for tables.
  /// 
  /// [paragraphReplacements] A map of placeholder keys to their replacement paragraph values
  Future<void> replaceParagraph(
      {required Map<String, dynamic> paragraphReplacements}) async {
    // Wait for the file to be extracted
    await isExtracted;

    // Read XML
    String documentFilePath = generateXMLPathDocument();

    // Replace placeholders
    _replacePlaceholdersParagraph(
        documentPath: documentFilePath, replacements: paragraphReplacements);
  }

  /// Replaces image placeholders with actual images.
  /// 
  /// [imageReplacements] A map of placeholder keys to MsImageUtil instances
  Future<void> replaceImage(
      {required Map<String, MsImageUtil> imageReplacements}) async {
    // Wait for the file to be extracted
    await isExtracted;

    // Read XML
    String documentFilePath = generateXMLPathDocument();

    // Replace placeholders
    await _replacePlaceholdersImage(
        documentPath: documentFilePath, replacements: imageReplacements);
  }

  /// Replaces images within text boxes in the document.
  /// 
  /// [imageReplacements] A map of placeholder keys to image files
  Future<void> replaceImageInTextBox(
      {required Map<String, File?> imageReplacements}) async {
    // Wait for the file to be extracted
    await isExtracted;

    // Read XML
    String documentFilePath = generateXMLPathDocument();

    // Replace placeholders
    await _replacePlaceholdersImageInTextBox(
        documentPath: documentFilePath, replacements: imageReplacements);
  }

  /// Adds an image relationship to the document and returns the relationship ID.
  /// 
  /// [image] The image file to be added
  /// [relationshipFilePath] Optional path to the relationship file
  /// Returns the new relationship ID (rId) for the added image
  Future<String> addImageRelationship(
      {required File image, String? relationshipFilePath}) async {
    // Extract file to tmp if not extracted
    await isExtracted;

    // Generate the XML path if not provided the default is the document XML path
    relationshipFilePath ??= generateXMLPathRelationshipByXML(generateXMLPathDocument());

    // Check if the image is valid
    if (!isImage(image.path)) {
      throw Exception("Invalid image format. Only PNG and JPG are allowed!");
    }

    // If the image extension is not png ms word gives error
    String imageExtension = ".png";

    // Read XML
    String xmlContent = getContent(relationshipFilePath);

    // Generate new Rid and new imageName
    int newRIdNumber = getNewRIdNumber(xmlContent, relationshipFilePath);
    String newRid = "rId$newRIdNumber";

    // Create image name based on the file type
    String imageName = "image$newRIdNumber$imageExtension";

    // Save the image with the correct name
    await saveImage(image, imageName);

    // Generate new relation
    String newRelation = replacePlaceholders(
        content: msRelationShipTemplate,
        replacements: {"imageIdNumber": newRIdNumber, "imageName": imageName});
    xmlContent = addRelationship(xmlContent, newRelation);

    // Update the document map
    xmlContentMap[relationshipFilePath] = xmlContent;

    return newRid;
  }

  /// Replaces text placeholders in document headers.
  /// 
  /// [textReplacements] A map of placeholder keys to their replacement values
  Future<void> replaceTextHeader(
      {required Map<String, dynamic> textReplacements}) async {
    // Wait for the file to be extracted
    await isExtracted;


    List<String> headerFilePathList = generateAllXMLPathHeaders();

    for (String headerFilePath in headerFilePathList) {
      _replacePlaceholders(
          documentPath: headerFilePath, replacements: textReplacements);
    }

  }

  /// Replaces paragraph placeholders in document headers.
  /// 
  /// [paragraphReplacements] A map of placeholder keys to their replacement paragraph values
  Future<void> replaceParagraphHeader(
      {required Map<String, dynamic> paragraphReplacements}) async {
    // Wait for the file to be extracted
    await isExtracted;

    List<String> headerFilePathList = generateAllXMLPathHeaders();

    for (String headerFilePath in headerFilePathList) {
      _replacePlaceholdersParagraph(
          documentPath: headerFilePath, replacements: paragraphReplacements);
    }
  }

  /// Replaces images in text boxes within document headers.
  /// 
  /// [imageReplacements] A map of placeholder keys to image files
  Future<void> replaceImageInTextBoxHeader(
      {required Map<String, File?> imageReplacements}) async {
    // Wait for the file to be extracted
    await isExtracted;

    List<String> headerFilePathList = generateAllXMLPathHeaders();

    for (String headerFilePath in headerFilePathList) {
      await _replacePlaceholdersImageInTextBox(
          documentPath: headerFilePath, replacements: imageReplacements);
    }
  }

  /// Replaces text placeholders in document footers.
  /// 
  /// [textReplacements] A map of placeholder keys to their replacement values
  Future<void> replaceTextFooter(
      {required Map<String, dynamic> textReplacements}) async {
    // Wait for the file to be extracted
    await isExtracted;

    // Read XML
    List<String> footerFilePathList = generateAllXMLPathFooters();
    for (String footerFilePath in footerFilePathList) {
      _replacePlaceholders(
          documentPath: footerFilePath, replacements: textReplacements);
    }
  }

  /// Replaces paragraph placeholders in document footers.
  /// 
  /// [paragraphReplacements] A map of placeholder keys to their replacement paragraph values
  Future<void> replaceParagraphFooter(
      {required Map<String, dynamic> paragraphReplacements}) async {
    // Wait for the file to be extracted
    await isExtracted;

    // Read XML
    List<String> footerFilePathList = generateAllXMLPathFooters();
    for (String footerFilePath in footerFilePathList) {
      _replacePlaceholdersParagraph(
          documentPath: footerFilePath, replacements: paragraphReplacements);
    }
  }

/*
  !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
  Replace Helper Functions
  !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
*/

  /// Replaces placeholders in a string with their corresponding values.
  /// 
  /// [content] The string containing placeholders
  /// [replacements] A map of placeholder keys to their replacement values
  /// Returns the string with all placeholders replaced
  static String replacePlaceholders(
      {required String content, required Map<String, dynamic> replacements}) {
    replacements.forEach((key, value) {
      RegExp regex = RegExp(r'{[^{}]*\b' + RegExp.escape(key) + r'\b[^{}]*}');
      if (content.contains(regex)) {
        content = content.replaceAllMapped(regex, (match) => value.toString());
      }
    });

    return content;
  }

  /// Internal method to replace placeholders in a document.
  /// 
  /// [documentPath] Path to the document XML file
  /// [replacements] A map of placeholder keys to their replacement values
  void _replacePlaceholders(
      {required String documentPath,
      required Map<String, dynamic> replacements}) {
    String content = getContent(documentPath);

    replacements.forEach((key, value) {
      RegExp regex = RegExp(r'{[^{}]*\b' + RegExp.escape(key) + r'\b[^{}]*}');
      if (content.contains(regex)) {
        content = content.replaceAllMapped(regex, (match) => value.toString());
      }
    });

    xmlContentMap[documentPath] = content;
  }

  /// Replaces paragraph placeholders in a string with their corresponding values.
  /// 
  /// [content] The string containing paragraph placeholders
  /// [replacements] A map of placeholder keys to their replacement paragraph values
  /// Returns the string with all paragraph placeholders replaced
  static String replacePlaceholdersParagraph(
      {required String content, required Map<String, dynamic> replacements}) {
    replacements.forEach((key, value) {
      RegExp regex = RegExp(
          r'<w:p\b[^>]*>(?:(?!<w:p\b).)*?\{[^{}]*\b' +
              RegExp.escape(key) +
              r'\b[^{}]*}.*?<\/w:p>',
          dotAll: true);

      if (content.contains(regex)) {
        content = content.replaceAllMapped(regex, (match) => value.toString());
      }
    });

    return content;
  }

  /// Internal method to replace paragraph placeholders in a document.
  /// 
  /// [documentPath] Path to the document XML file
  /// [replacements] A map of placeholder keys to their replacement paragraph values
  void _replacePlaceholdersParagraph(
      {required String documentPath,
      required Map<String, dynamic> replacements}) {
    String content = getContent(documentPath);

    replacements.forEach((key, value) {
      RegExp regex = RegExp(
          r'<w:p\b[^>]*>(?:(?!<w:p\b).)*?\{[^{}]*\b' +
              RegExp.escape(key) +
              r'\b[^{}]*}.*?<\/w:p>',
          dotAll: true);

      if (content.contains(regex)) {
        content = content.replaceAllMapped(regex, (match) => value.toString());
      }
    });

    xmlContentMap[documentPath] = content;
  }

  /// Internal method to replace image placeholders in a document.
  /// 
  /// [documentPath] Path to the document XML file
  /// [replacements] A map of placeholder keys to MsImageUtil instances
  Future<void> _replacePlaceholdersImage(
      {required String documentPath,
      required Map<String, MsImageUtil> replacements}) async {
    String content = getContent(documentPath);

    replacements.forEach((key, value) async {
      // Regex to match a paragraph containing the placeholder { .* key .* }
      RegExp regex = RegExp(
          r'<w:p\b[^>]*>(?:(?!<w:p\b).)*?\{[^{}]*\b' +
              RegExp.escape(key) +
              r'\b[^{}]*}.*?<\/w:p>',
          dotAll: true);

      if (content.contains(regex)) {
        // add the image to the relationship
        String newRid = await addImageRelationship(
            image: value.image, relationshipFilePath: generateXMLPathRelationshipByXML(documentPath));

        // generate image template
        String updatedImageTemplate = replacePlaceholders(
            content: msImageTemplate,
            replacements: {
              "width": value.width,
              "height": value.height,
              "imageId": newRid
            });

        // Replace the entire paragraph with the replacement value
        content = content.replaceAllMapped(
            regex, (match) => updatedImageTemplate.toString());
      }
    });

    xmlContentMap[documentPath] = content;
  }

  /// Internal method to replace images in text boxes within a document.
  /// 
  /// [documentPath] Path to the document XML file
  /// [replacements] A map of placeholder keys to image files
  /// [padding] Padding value for the text box (default: 14000)
  Future<void> _replacePlaceholdersImageInTextBox(
      {required String documentPath,
      required Map<String, File?> replacements,
      int padding = 14000}) async {
    String content = getContent(documentPath);

    for (var entry in replacements.entries) {
      String key = entry.key;
      File? value = entry.value;

      if (value == null) continue;

      String startOfMcChoice = r'<mc:Choice\b[^>]*>(?:(?!<mc:Choice\b).)*?';
      String dimensionInfo =
          r'<wp:extent\s*?cx="(\d+)"\s*?cy="(\d+)"\s*?/>(?:(?!<mc:Choice\b).)*?';
      String placeHolder = r'(<w:p\b[^>]*>(?:(?!<w:p\b).)*?\{[^{}]*\b' +
          RegExp.escape(key) +
          r'\b[^{}]*}.*?<\/w:p>)';
      String endOfMcChoice = r'(?:(?!<mc:Choice\b).)*?</mc:Choice>';

      RegExp regex = RegExp(
          startOfMcChoice + dimensionInfo + placeHolder + endOfMcChoice,
          dotAll: true);

      if (content.contains(regex)) {
        MsImageUtil msImage = await MsImageUtil.createInstance(image: value);
        String newRid = await addImageRelationship(
            image: msImage.image, relationshipFilePath: generateXMLPathRelationshipByXML(documentPath));

        // Replace the content correctly
        content = content.replaceAllMapped(regex, (match) {
          int wantedWidth = int.parse(match.group(1)!);
          int wantedHeight = int.parse(match.group(2)!);
          msImage.wantedWidth = wantedWidth - padding;
          msImage.wantedHeight = wantedHeight - padding;

          // Generate updated image template
          String updatedImageTemplate =
              replacePlaceholders(content: msImageTemplate, replacements: {
            "width": msImage.width,
            "height": msImage.height,
            "imageId": newRid,
          });

          // Replace the third group with the updated image template
          String fullMatch = match.group(0)!;
          String thirdGroup = match.group(3)!;
          String replaced =
              fullMatch.replaceFirst(thirdGroup, updatedImageTemplate);

          return replaced;
        });
      }
    }

    xmlContentMap[documentPath] = content;
  }

/*
  !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
  File Helper Functions
  !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
*/

  /// Extracts a DOCX file to a temporary directory.
  /// 
  /// [filePath] Path to the DOCX file
  /// [isAsset] Whether the file is an asset
  /// Returns true if extraction was successful
  Future<bool> extractDocx(String filePath, bool isAsset) async {
    // Load Document
    ByteData? bytes =
        isAsset ? await loadDocsAsset(filePath) : await loadFile(filePath);

    // Create Tmp Path
    String tmpPath = await createTmpDir();

    // Unzip the docx file
    final archive = ZipDecoder().decodeBytes(bytes.buffer.asUint8List());

    // Extract the files
    for (final file in archive.files) {
      final filePath = path.join(tmpPath, file.name);

      if (file.isFile) {
        final newFile = File(filePath);
        newFile.createSync(recursive: true);
        newFile.writeAsBytesSync(file.content);
      }
    }

    // Set the tempPath
    this.tmpPath = tmpPath;
    return true;
  }

  /// Saves the modified document as a new DOCX file.
  Future<void> saveDocx() async {
    final tempDirFolder = Directory(tmpPath!);

    // Save the changed documents
    xmlContentMap.forEach((filePathKey, xmlContentValue) {
      writeContent(filePathKey, xmlContentValue);
    });

    // Recompress the docx file
    final newArchive = Archive();
    for (final file in tempDirFolder.listSync(recursive: true)) {
      final relativePath = path.relative(file.path,
          from: tmpPath); // platform-safe relative path

      if (file is File) {
        final fileBytes = file.readAsBytesSync();
        newArchive
            .addFile(ArchiveFile(relativePath, fileBytes.length, fileBytes));
      }
    }

    // Write the newArchive to a file or handle it as needed
    final compressedFile = File(savePath);
    List<int>? encodedData = ZipEncoder().encode(newArchive);

    if (encodedData == null) {
      throw Exception("Failed To Encode Data");
    }

    // Write the data to compressed file
    compressedFile.writeAsBytesSync(encodedData);

    // Delete the modified document and extract again
    // dispose();
    // await extractDocx(filePath, isAsset);
  }

  /// Creates a temporary directory for document manipulation.
  /// 
  /// Returns the path to the created temporary directory
  Future<String> createTmpDir() async {
    // Use a writable temporary directory
    final tempDir = await getTemporaryDirectory();

    final fileName = path.basename(filePath);
    final timestamp = DateTime.now().millisecondsSinceEpoch;
    final tmpDirName = '${fileName}_$timestamp';
    final tmpPath = path.join(tempDir.path, tmpDirName);
    final tempDirFolder = Directory(tmpPath);

    // Create the directory if it doesn't exist
    if (!await tempDirFolder.exists()) {
      await tempDirFolder.create(recursive: true);
    }

    return tmpPath;
  }

  /// Loads a document from the assets bundle.
  /// 
  /// [filename] Name of the file in the assets directory
  /// Returns the file contents as ByteData
  Future<ByteData> loadDocsAsset(String filename) async {
    // Create filePath
    final assetPath = 'assets/docs/$filename';

    // Check if the file exists in assets
    try {
      await rootBundle.load(assetPath);
    } catch (e) {
      throw Exception("File $assetPath not found in assets.");
    }

    final bytes = await rootBundle.load(assetPath);
    return bytes;
  }

  /// Loads a file from the local filesystem.
  /// 
  /// [filePath] Path to the file
  /// Returns the file contents as ByteData
  Future<ByteData> loadFile(String filePath) async {
    try {
      // Check if the file exists at the given path
      final file = File(filePath);
      if (!await file.exists()) {
        throw Exception("File $filePath not found.");
      }

      // Read the file as bytes
      final bytes = await file.readAsBytes();

      // Return the bytes as ByteData
      return ByteData.sublistView(Uint8List.fromList(bytes));
    } catch (e) {
      throw Exception("Error loading file: $e");
    }
  }

  /// Gets the content of an XML file.
  /// 
  /// [filePath] Path to the XML file
  /// Returns the file contents as a string
  String getContent(String filePath) {

    // If xml content is reached before return it
    if (xmlContentMap[filePath] != null) {
      return xmlContentMap[filePath]!;
    }

    final xmlFile = File(filePath);

    if (!xmlFile.existsSync()) {
      throw Exception("$filePath not found");
    }

    String xmlContent = xmlFile.readAsStringSync();
    xmlContentMap[filePath] = xmlContent;
    return xmlContent;
  }

  /// Gets a File object for the specified path.
  /// 
  /// [filePath] Path to the file
  /// Returns a File object
  File getFile(String filePath) {
    final documentXmlFile = File(filePath);
    if (!documentXmlFile.existsSync()) {
      throw Exception("document.xml not found");
    }
    return documentXmlFile;
  }

  /// Writes content to an XML file.
  /// 
  /// [filePath] Path to the XML file
  /// [xmlContent] Content to write to the file
  void writeContent(String filePath, String xmlContent) {
    final documentXmlFile = File(filePath);

    if (!documentXmlFile.existsSync()) {
      throw Exception("document.xml not found");
    }

    documentXmlFile.writeAsStringSync(xmlContent);
  }

  /// Saves an image file to the document's media directory.
  /// 
  /// [image] The image file to save
  /// [imageName] Name to save the image as
  Future<void> saveImage(File image, String imageName) async {
    try {
      final String imagePath = path.join(generatePathMedia(), imageName);
      await image.copy(imagePath);
    } catch (e) {
      throw Exception("Error saving image: $e");
    }
  }

  /// Adds a relationship entry to an XML file.
  /// 
  /// [xmlContent] The XML content to modify
  /// [relationshipContent] The relationship entry to add
  /// Returns the modified XML content
  String addRelationship(String xmlContent, String relationshipContent) {
    // Parse the XML content
    final document = xml.XmlDocument.parse(xmlContent);

    // Parse the relationship content
    final relationshipElement =
        xml.XmlDocument.parse(relationshipContent).rootElement;

    // Find the <Relationships> element in the XML content
    final relationshipsElement = document.findElements('Relationships').first;

    // Check if the relationshipElement already has a parent and remove it if needed
    if (relationshipElement.parent != null) {
      relationshipElement.parent!.children.remove(relationshipElement);
    }

    // Add the new relationship content into the <Relationships> element
    relationshipsElement.children.add(relationshipElement);
  
    // Return the updated XML as a string
    return document.toXmlString(pretty: true, indent: '  ');
  }

  /// Gets a new relationship ID number for the specified file.
  /// 
  /// [xmlContent] The XML content to analyze
  /// [filePath] Path to the relationship file
  /// Returns the next available relationship ID number
  int getNewRIdNumber(String xmlContent, String filePath) {
    currentRIdNumberMap[filePath] ??= 0;

    // Give the current rid number +1 if possible
    if (currentRIdNumberMap[filePath] != 0) {
      currentRIdNumberMap[filePath] = currentRIdNumberMap[filePath]! + 1;
      return currentRIdNumberMap[filePath]!;
    }

    // Define the regular expression to match rId attributes
    // Find all matches in the xmlContent

    final regExp = RegExp(r'rId(\d+)');
    final matches = regExp.allMatches(xmlContent);

    // If no matches are found, return 1
    if (matches.isEmpty) {
      currentRIdNumberMap[filePath] = 1;
      return currentRIdNumberMap[filePath]!;
    }

    // current rid number is 0
    // Loop through each match and check if the current rId number is greater than the last one
    for (final match in matches) {
      final rIdNumber = int.parse(match.group(1)!);

      // If current rId is greater than the last one, update lastRIdNumber
      currentRIdNumberMap[filePath] = rIdNumber > currentRIdNumberMap[filePath]!
          ? rIdNumber
          : currentRIdNumberMap[filePath]!;
    }

    currentRIdNumberMap[filePath] = currentRIdNumberMap[filePath]! + 1;
    return currentRIdNumberMap[filePath]!;
  }

/*
  !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
  MS Paths Helper Functions
  !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
*/

  /// Generates the path to the document.xml file.
  /// 
  /// Returns the path to document.xml
  String generateXMLPathDocument() {
    return path.joinAll([tmpPath!, "word", "document.xml"]);
  }

  /// Generates paths to all header XML files in the document.
  /// 
  /// Returns a list of paths to header XML files
  List<String> generateAllXMLPathHeaders() {
    List<String> headers = [];
    String wordDirPath = path.joinAll([tmpPath!, "word"]);
    Directory wordDir = Directory(wordDirPath);

    if (wordDir.existsSync()) {
      List<FileSystemEntity> files = wordDir.listSync();
      for (var file in files) {
        if (file is File && path.basename(file.path).startsWith('header')) {
          headers.add(file.path);
        }
      }
    }
    return headers;
  }

  /// Generates paths to all footer XML files in the document.
  /// 
  /// Returns a list of paths to footer XML files
  List<String> generateAllXMLPathFooters() {
    String wordDirPath = path.joinAll([tmpPath!, "word"]);
    Directory wordDir = Directory(wordDirPath);
    List<String> footers = [];

    if (wordDir.existsSync()) {
      List<FileSystemEntity> files = wordDir.listSync();
      for (var file in files) {
        if (file is File && path.basename(file.path).startsWith('footer')) {
          footers.add(file.path);
        }
      }
    }

    return footers;
  }

  /// Generates the path to a relationship file based on an XML file path.
  /// 
  /// [xmlPath] Path to the XML file
  /// Returns the path to the corresponding relationship file
  String generateXMLPathRelationshipByXML(String xmlPath) {
    String relPath = path
        .joinAll([tmpPath!, "word", "_rels", "${path.basename(xmlPath)}.rels"]);

    // Create _rels directory if it doesn't exist
    Directory relsDir = Directory(path.dirname(relPath));
    if (!relsDir.existsSync()) {
      relsDir.createSync(recursive: true);
    }

    // Create relationship file if it doesn't exist
    File relFile = File(relPath);
    if (!relFile.existsSync()) {
      relFile.writeAsStringSync(
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>');
    }

    return relPath;
  }

  /// Generates the path to the document's media directory.
  /// 
  /// Returns the path to the media directory
  String generatePathMedia() {
    String mediaPath = path.joinAll([tmpPath!, "word", "media"]);
    Directory mediaDir = Directory(mediaPath);
    if (!mediaDir.existsSync()) {
      mediaDir.createSync(recursive: true);
    }
    return mediaPath;
  }

/*
  !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
  User Functions
  !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
*/

  /// Creates a text box containing an image.
  /// 
  /// [textBoxWidth] Width of the text box in EMU units
  /// [textBoxHeight] Height of the text box in EMU units
  /// [image] The image file to insert
  /// [padding] Padding value for the text box (default: 14000)
  /// Returns the XML content for the text box with image
  Future<String> createTextBoxImage(
      int textBoxWidth, int textBoxHeight, File image,
      {int padding = 14000}) async {
    if (image.path == "") {
      return "";
    }

    // Apply padding to the text box width and height
    int wantedWidth = textBoxWidth - padding;
    int wantedHeight = textBoxHeight - padding;

    String imageId = await addImageRelationship(image: image);
    MsImageUtil msImageUtil = await MsImageUtil.createInstance(
        image: image, wantedHeight: wantedHeight, wantedWidth: wantedWidth);

    Map<String, dynamic> imageReplacements = {
      "textBoxWidth": textBoxWidth,
      "textBoxHeight": textBoxHeight,
      "imageWidth": msImageUtil.width,
      "imageHeight": msImageUtil.height,
      "imageId": imageId,
    };

    return replacePlaceholders(
        content: msTextBoxImageTemplate, replacements: imageReplacements);
  }

  /// Creates multiple text boxes containing images.
  /// 
  /// [textBoxWidthCm] Width of text boxes in centimeters
  /// [textBoxHeightCm] Height of text boxes in centimeters
  /// [firstImageWidthCm] Width of first image in centimeters (optional)
  /// [firstImageHeightCm] Height of first image in centimeters (optional)
  /// [imageFiles] List of image files to insert
  /// Returns the XML content for all text boxes with images
  Future<String> createMultiTextBoxImages({
    required double textBoxWidthCm,
    required double textBoxHeightCm,
    double? firstImageWidthCm,
    double? firstImageHeightCm,
    required List<File> imageFiles,
  }) async {
    String images = "";

    firstImageWidthCm ??= textBoxWidthCm;
    firstImageHeightCm ??= textBoxHeightCm;

    for (int i = 0; i < imageFiles.length; i++) {
      var image = imageFiles[i];
      if (image.path == "") continue;

      int textBoxHeightEmu = DocxUtil.cmToEmu(textBoxHeightCm);
      int textBoxWidthEmu = DocxUtil.cmToEmu(textBoxWidthCm);

      if (i == 0) {
        textBoxHeightEmu = DocxUtil.cmToEmu(firstImageHeightCm);
        textBoxWidthEmu = DocxUtil.cmToEmu(firstImageWidthCm);
      }

      images +=
          await createTextBoxImage(textBoxWidthEmu, textBoxHeightEmu, image);
      images += msPageBreak;
    }

    return images;
  }

  /// Creates text boxes containing images with a title.
  /// 
  /// [title] The title text
  /// [titleFontSizeHalfPoints] Font size of the title in half-points (default: 22)
  /// [titleBold] Whether the title should be bold (default: false)
  /// [titleItalic] Whether the title should be italic (default: false)
  /// [titleUnderline] Whether the title should be underlined (default: false)
  /// [titleFont] Font family for the title (default: "Arial")
  /// [textBoxWidthCm] Width of text boxes in centimeters
  /// [textBoxHeightCm] Height of text boxes in centimeters
  /// [firstImageWidthCm] Width of first image in centimeters (optional)
  /// [firstImageHeightCm] Height of first image in centimeters (optional)
  /// [imageFiles] List of image files to insert
  /// Returns the XML content for the title and text boxes with images
  Future<String> createTextBoxImagesWithTitle({
    required String title,
    int titleFontSizeHalfPoints = 22,
    bool titleBold = false,
    bool titleItalic = false,
    bool titleUnderline = false,
    String titleFont = "Arial",
    required double textBoxWidthCm,
    required double textBoxHeightCm,
    double? firstImageWidthCm,
    double? firstImageHeightCm,
    required List<File> imageFiles,
  }) async {
    String images = "";

    firstImageWidthCm ??= textBoxWidthCm;
    firstImageHeightCm ??= textBoxHeightCm;

    images += createTextParagraph(
        text: title,
        fontSizeHalfPoints: titleFontSizeHalfPoints,
        bold: titleBold,
        italic: titleItalic,
        underline: titleUnderline,
        font: titleFont);

    for (int i = 0; i < imageFiles.length; i++) {
      var image = imageFiles[i];
      if (image.path == "") continue;
      int textBoxHeightEmu = DocxUtil.cmToEmu(textBoxHeightCm);
      int textBoxWidthEmu = DocxUtil.cmToEmu(textBoxWidthCm);

      if (i == 0) {
        textBoxHeightEmu = DocxUtil.cmToEmu(firstImageHeightCm);
        textBoxWidthEmu = DocxUtil.cmToEmu(firstImageWidthCm);
      }

      images +=
          await createTextBoxImage(textBoxWidthEmu, textBoxHeightEmu, image);
      images += msPageBreak;
    }

    return images;
  }

  /// Creates a text paragraph with specified formatting.
  /// 
  /// [text] The text content
  /// [font] Font family (default: "Arial")
  /// [fontSizeHalfPoints] Font size in half-points (default: 22)
  /// [bold] Whether the text should be bold (default: false)
  /// [italic] Whether the text should be italic (default: false)
  /// [underline] Whether the text should be underlined (default: false)
  /// [alignment] Text alignment (default: "left")
  /// Returns the XML content for the formatted text paragraph
  String createTextParagraph({
    required String text,
    String font = "Arial",
    int fontSizeHalfPoints = 22,
    bool bold = false,
    bool italic = false,
    bool underline = false,
    String alignment = "left",
  }) {
    return msTextTemplate
        .replaceAll("{font}", font)
        .replaceAll("{fontSize}", fontSizeHalfPoints.toString())
        .replaceAll("{boldTag}", bold ? "<w:b />" : "")
        .replaceAll("{italicTag}", italic ? "<w:i />" : "")
        .replaceAll(
            "{underlineTag}", underline ? "<w:u w:val=\"single\" />" : "")
        .replaceAll("{alignment}", alignment)
        .replaceAll("{insertText}", text);
  }

  /// Converts centimeters to EMU (English Metric Units).
  /// 
  /// [cm] Length in centimeters
  /// Returns the length in EMU units
  static int cmToEmu(double cm) {
    return (cm * 360000).round();
  }

  /// Checks if a file is an image file.
  /// 
  /// [filePath] Path to the file to check
  /// Returns true if the file is a PNG, JPG, or JPEG image
  static bool isImage(String filePath) {
    // check if the file is an image it can be png, jpg, jpeg
    return filePath.toLowerCase().endsWith('.png') ||
        filePath.toLowerCase().endsWith('.jpg') ||
        filePath.toLowerCase().endsWith('.jpeg');
  }
}
