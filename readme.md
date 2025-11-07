# Flutter MS Word Library

A powerful Flutter library for programmatically creating and manipulating Microsoft Word (.docx) documents. This library provides comprehensive functionality for text replacement, image insertion, table manipulation, and complex document generation.

[![Dart](https://img.shields.io/badge/Dart-0175C2?logo=dart&logoColor=white)](https://dart.dev)
[![Flutter](https://img.shields.io/badge/Flutter-02569B?logo=flutter&logoColor=white)](https://flutter.dev)

## Features

✨ **Text Manipulation**
- Replace text placeholders with dynamic values using `{key}` format
- Replace entire paragraphs for complex content
- Support for headers and footers

🖼️ **Image Handling**
- Insert images with automatic aspect ratio management
- Replace image placeholders in text boxes
- Support for PNG and JPG formats
- Automatic dimension calculation and scaling

📊 **Table Support**
- Generate dynamic tables with custom rows
- Pre-built templates for personnel, projects, and transcripts
- Easy table row insertion and manipulation

📄 **Document Structure**
- Work with headers and footers independently
- Manage multiple document sections
- Create complex document layouts with text boxes

🔧 **Advanced Features**
- Extract and manipulate DOCX XML structure
- Add image relationships dynamically
- Support for both asset and file system documents
- Automatic temporary file management

## Placeholder Format

The library uses a simple and consistent placeholder format throughout all operations:

**Format**: `{key}`

Where `key` is the identifier you use in your code to replace the placeholder.

**Examples**:
- `{companyName}` → Replaced with company name value
- `{date}` → Replaced with date value
- `{employeeList}` → Replaced with employee list content
- `{logoImage}` → Replaced with an image

**In Your Template** (Word document):
```
Dear {customerName},

Your order #{orderNumber} has been confirmed on {orderDate}.

Thank you for your business!
```

**In Your Code**:
```dart
await docxUtil.replaceText(
  textReplacements: {
    'customerName': 'John Doe',
    'orderNumber': '12345',
    'orderDate': '2025-11-07',
  }
);
```

**Result**:
```
Dear John Doe,

Your order #12345 has been confirmed on 2025-11-07.

Thank you for your business!
```

## Architecture Overview

### Core Classes

#### 1. `DocxUtil` - Main Document Manipulation Class

The heart of the library, `DocxUtil` provides a comprehensive API for working with DOCX files.

**Design Pattern**: Builder/Facade Pattern
- Simplifies complex XML manipulation behind an easy-to-use interface
- Handles DOCX file lifecycle (extract → modify → save)

**Key Responsibilities**:
- DOCX file extraction and compression
- XML content manipulation
- Temporary file management
- Relationship management for embedded resources

#### 2. `MsImageUtil` - Image Dimension Management

A utility class that handles image dimension calculations while maintaining aspect ratios.

**Key Features**:
- Automatic original dimension detection
- Aspect ratio preservation
- EMU (English Metric Units) conversion
- Support for wanted width/height constraints

#### 3. `CreateDocumentUtil` - Document Generation Helper

A high-level utility that demonstrates how to use `DocxUtil` for generating complete documents from templates.

**Purpose**: Provides ready-to-use methods for generating specific document types (KGB, EK-series documents, etc.)

## How DocxUtil Works

### 1. Initialization & Extraction

```dart
DocxUtil docxUtil = DocxUtil(
  filePath: 'template.docx',    // Source template
  savePath: '/path/to/output.docx',  // Output path
  isAsset: true                  // Load from assets or file system
);
```

**Internal Process**:
1. Validates the file has `.docx` extension
2. Loads the file (from assets bundle or file system)
3. Extracts the ZIP archive to a temporary directory
4. Stores the temporary path for XML manipulation

### 2. Content Manipulation

The library uses XML manipulation to modify document content:

```dart
// Text Replacement
await docxUtil.replaceText(
  textReplacements: {
    'companyName': 'ACME Corp',
    'date': '2025-11-07',
  }
);

// Paragraph Replacement (for tables/complex content)
await docxUtil.replaceParagraph(
  paragraphReplacements: {
    'tableContent': generatedTableXML,
  }
);

// Image Replacement
await docxUtil.replaceImage(
  imageReplacements: {
    'logoImage': MsImageUtil.createInstance(
      image: logoFile,
      wantedWidth: 300,
      wantedHeight: 200,
    ),
  }
);
```

**XML Manipulation Strategy**:
- Uses regular expressions to find placeholders in `{key}` format (e.g., `{companyName}`)
- Placeholders are case-sensitive and must match exactly
- Maintains XML structure integrity
- Caches modified XML content in memory (`xmlContentMap`)
- Only writes changes when `saveDocx()` is called

**Note**: The placeholder format is always `{key}` where the key matches your replacement map keys.

### 3. Headers & Footers

The library provides dedicated methods for header/footer manipulation:

```dart
// Replace text in headers
await docxUtil.replaceTextHeader(
  textReplacements: {'headerTitle': 'Document Title'}
);

// Replace text in footers
await docxUtil.replaceTextFooter(
  textReplacements: {'footerText': 'Page Footer'}
);

// Replace images in header text boxes
await docxUtil.replaceImageInTextBoxHeader(
  imageReplacements: {'headerLogo': logoFile}
);
```

**Implementation Details**:
- Automatically discovers all header files (`header1.xml`, `header2.xml`, etc.)
- Automatically discovers all footer files (`footer1.xml`, `footer2.xml`, etc.)
- Applies replacements to all discovered files

### 4. Image Management

#### Image Relationship System

DOCX files use a relationship system to link images. The library handles this automatically:

```dart
String newRid = await docxUtil.addImageRelationship(
  image: imageFile,
  relationshipFilePath: 'word/_rels/document.xml.rels'
);
// Returns: 'rId7' (relationship ID)
```

**Process**:
1. Determines the next available relationship ID
2. Copies the image to the document's `media` folder
3. Creates a relationship entry in the `.rels` file
4. Returns the relationship ID for XML reference

#### Image in Text Boxes

For images within text boxes (common in Word templates):

```dart
await docxUtil.replaceImageInTextBox(
  imageReplacements: {
    'logoPlaceholder': logoFile,
  },
  padding: 14000  // EMU padding
);
```

**Intelligent Processing**:
1. Locates the text box containing the placeholder
2. Extracts the text box dimensions
3. Calculates optimal image size (with padding)
4. Maintains aspect ratio
5. Generates proper XML with image reference

### 5. Advanced Features

#### Creating Text Boxes with Images

```dart
String textBoxXML = await docxUtil.createTextBoxImage(
  3000000,  // Width in EMU
  2000000,  // Height in EMU
  imageFile,
  padding: 14000
);
```

#### Creating Formatted Text

```dart
String textXML = docxUtil.createTextParagraph(
  text: 'Important Notice',
  font: 'Arial',
  fontSizeHalfPoints: 28,  // 14pt (half-points)
  bold: true,
  italic: false,
  underline: true,
  alignment: 'center',
);
```

#### Multiple Images in Text Boxes

```dart
String imagesXML = await docxUtil.createMultiTextBoxImages(
  textBoxWidthCm: 15.0,
  textBoxHeightCm: 10.0,
  firstImageWidthCm: 17.0,
  firstImageHeightCm: 12.0,
  imageFiles: [image1, image2, image3],
);
```

### 6. Document Finalization

```dart
// Save the modified document
await docxUtil.saveDocx();

// Clean up temporary files
docxUtil.dispose();
```

**Save Process**:
1. Writes all cached XML content to their respective files
2. Recursively compresses the temporary directory back into a ZIP
3. Saves the archive with `.docx` extension
4. (Optional) Cleans up temporary files with `dispose()`

## Design Patterns & Principles

### 1. **Template Method Pattern**

The library uses template DOCX files as starting points, allowing for consistent document structure while enabling dynamic content.

### 2. **Lazy Loading & Caching**

- XML content is loaded only when needed
- Modified content is cached in `xmlContentMap`
- All changes are written atomically during save

### 3. **Separation of Concerns**

- **DocxUtil**: Core document manipulation
- **MsImageUtil**: Image dimension calculations
- **CreateDocumentUtil**: High-level document generation
- **Constants**: XML templates and configurations

### 4. **Resource Management**

```dart
// Explicit resource cleanup
void dispose() {
  final tempDirFolder = Directory(tmpPath!);
  tempDirFolder.deleteSync(recursive: true);
  xmlContentMap.clear();
}
```

Follows Flutter's disposal pattern for cleanup of temporary resources.

### 5. **Type Safety with Generics**

```dart
Future<void> replaceText({
  required Map<String, dynamic> textReplacements
}) async { ... }

Future<void> replaceImage({
  required Map<String, MsImageUtil> imageReplacements
}) async { ... }
```

## XML Templates

The library includes pre-built XML templates for common Word elements:

### Text & Formatting
- `msTextTemplate` - Formatted text paragraphs
- `msPageBreak` - Page break
- `msLineBreak` - Line break

### Images
- `msImageTemplate` - Inline images
- `msTextBoxImageTemplate` - Images in text boxes
- `msRelationShipTemplate` - Image relationship entries

### Tables
- `permissionGivenPersonalTable` - Personnel permission table
- `shareHolderPersonalTable` - Shareholder information table
- `authorizedPersonalTable` - Authorized personnel table
- `projectTable` - Project listing table
- `transcriptTable` - Document transcript table

### Table Rows
- `permissionGivenPersonalRow`
- `shareHolderPersonalRow`
- `authorizedPersonalRow`
- `projectRow`
- `transcriptRow`

### Plans & Images
- `planImageTemplate` - Plan/diagram image layout
- `planImageTemplateTogek` - TOGEK-specific plan layout

## Constants & Configuration

### File Naming

Predefined save file names for generated documents:

```dart
const String saveKgbFileName = "KGB Üst Yazısı.docx";
const String saveEk0FileName = "Ek-0 Başvuru Yazısı.docx";
// ... more predefined names
```

### Keys

Predefined placeholder keys:

```dart
const msPermissionGivenPeronalListKey = "permissionGivenPersonalList";
const msShareHolderPersonalKey = "shareHolderPersonalList";
const msAuthorizedPersonalKey = "authorizedPersonalList";
// ... more keys
```

## Usage Examples

### Example 1: Simple Text Replacement

```dart
DocxUtil docxUtil = DocxUtil(
  filePath: 'letter_template.docx',
  savePath: '/storage/letter_output.docx',
  isAsset: true,
);

await docxUtil.replaceText(
  textReplacements: {
    'recipientName': 'John Doe',
    'date': '2025-11-07',
    'subject': 'Project Proposal',
  }
);

await docxUtil.saveDocx();
docxUtil.dispose();
```

### Example 2: Document with Logo

```dart
DocxUtil docxUtil = DocxUtil(
  filePath: 'company_report.docx',
  savePath: '/storage/report_output.docx',
  isAsset: true,
);

// Replace header logo
await docxUtil.replaceImageInTextBoxHeader(
  imageReplacements: {
    'companyLogo': File('/path/to/logo.png'),
  }
);

// Replace body text
await docxUtil.replaceText(
  textReplacements: {
    'reportTitle': 'Q4 Financial Report',
    'year': '2025',
  }
);

await docxUtil.saveDocx();
docxUtil.dispose();
```

### Example 3: Dynamic Table Generation

```dart
DocxUtil docxUtil = DocxUtil(
  filePath: 'employee_list.docx',
  savePath: '/storage/employees.docx',
  isAsset: true,
);

// Generate table rows
String tableRows = '';
for (var employee in employees) {
  tableRows += permissionGivenPersonalRow
      .replaceAll('{name}', employee.name)
      .replaceAll('{idNumber}', employee.id)
      .replaceAll('{duty}', employee.position);
}

// Replace table placeholder
String completeTable = permissionGivenPersonalTable
    .replaceAll('{insert_row}', tableRows);

await docxUtil.replaceParagraph(
  paragraphReplacements: {
    'employeeTable': completeTable,
  }
);

await docxUtil.saveDocx();
docxUtil.dispose();
```

### Example 4: Complex Document with Multiple Images

```dart
DocxUtil docxUtil = DocxUtil(
  filePath: 'project_report.docx',
  savePath: '/storage/project_output.docx',
  isAsset: true,
);

// Create image pages
String imagePages = await docxUtil.createMultiTextBoxImages(
  textBoxWidthCm: 15.0,
  textBoxHeightCm: 10.0,
  imageFiles: [
    File('photo1.jpg'),
    File('photo2.jpg'),
    File('photo3.jpg'),
  ],
);

// Add title and images
String content = await docxUtil.createTextBoxImagesWithTitle(
  title: 'Project Progress Photos',
  titleFontSizeHalfPoints: 28,
  titleBold: true,
  textBoxWidthCm: 15.0,
  textBoxHeightCm: 10.0,
  imageFiles: [File('photo1.jpg'), File('photo2.jpg')],
);

await docxUtil.replaceParagraph(
  paragraphReplacements: {
    'photoSection': content,
  }
);

await docxUtil.saveDocx();
docxUtil.dispose();
```

## Dependencies

The library relies on the following Flutter packages:

```yaml
dependencies:
  flutter:
    sdk: flutter
  archive: ^3.0.0  # ZIP compression/decompression
  path: ^1.8.0     # Path manipulation
  path_provider: ^2.0.0  # Temporary directory access
  xml: ^6.0.0      # XML parsing and manipulation
  image: ^4.0.0    # Image dimension detection
```

## Utility Methods

### Unit Conversion

```dart
// Convert centimeters to EMU (English Metric Units)
int emu = DocxUtil.cmToEmu(10.5);  // 3780000 EMU
```

**Conversion Formula**: `EMU = CM × 360,000`

### Image Validation

```dart
bool isValid = DocxUtil.isImage(filePath);
// Returns true for .png, .jpg, .jpeg
```

## Performance Considerations

### 1. **Temporary File Management**

- Each `DocxUtil` instance creates a unique temporary directory
- Directory name includes timestamp to prevent conflicts
- **Important**: Always call `dispose()` to clean up temporary files

### 2. **Memory Optimization**

- XML content is loaded lazily
- Modified content is cached to avoid redundant file I/O
- Images are referenced, not embedded multiple times

### 3. **Asynchronous Operations**

All I/O operations are asynchronous to prevent blocking the UI thread:

```dart
await docxUtil.replaceText(...);      // Async
await docxUtil.replaceImage(...);     // Async
await docxUtil.saveDocx();            // Async
docxUtil.dispose();                    // Sync (cleanup)
```

## Best Practices

### 1. Resource Cleanup

Always dispose of `DocxUtil` instances:

```dart
DocxUtil? docxUtil;
try {
  docxUtil = DocxUtil(...);
  await docxUtil.replaceText(...);
  await docxUtil.saveDocx();
} finally {
  docxUtil?.dispose();
}
```

### 2. Template Organization

Keep template DOCX files in the `assets/docs/` directory:

```yaml
flutter:
  assets:
    - assets/docs/
```

### 3. Placeholder Format & Naming Convention

**Always use the format**: `{key}`

Use descriptive, camelCase keys:

**In Word Template**:
```
{companyName}
{employeeList}
{reportDate}
```

**In Code**:
```dart
await docxUtil.replaceText(
  textReplacements: {
    'companyName': 'ACME Corp',      // Note: No curly braces in code
    'employeeList': tableContent,
    'reportDate': '2025-11-07',
  }
);
```

**Important**: 
- Placeholders in Word documents include curly braces: `{key}`
- Keys in code do NOT include curly braces: `'key'`
- Placeholders are case-sensitive

### 4. Error Handling

```dart
try {
  DocxUtil docxUtil = DocxUtil(...);
  await docxUtil.replaceText(...);
  await docxUtil.saveDocx();
  return [true, "Document created successfully"];
} catch (e) {
  return [false, e.toString()];
} finally {
  docxUtil?.dispose();
}
```

### 5. Image Format

- Prefer PNG format for transparency support
- Ensure images are reasonably sized (not too large)
- Use `MsImageUtil` for automatic dimension management

## Advanced Topics

### XML Structure Understanding

DOCX files are ZIP archives containing XML files:

```
document.docx
├── word/
│   ├── document.xml        # Main document content
│   ├── header1.xml         # Header content
│   ├── footer1.xml         # Footer content
│   ├── _rels/
│   │   └── document.xml.rels  # Relationships
│   └── media/
│       └── image1.png      # Embedded images
├── docProps/
│   └── core.xml            # Document properties
└── [Content_Types].xml     # MIME types
```

### Custom XML Templates

You can create custom XML templates by:

1. Creating a Word document with the desired layout
2. Saving as `.docx`
3. Extracting the ZIP
4. Copying the relevant XML structure
5. Adding placeholders using `{keyName}` format

**Placeholder Format in Custom Templates**:

When creating custom templates, always use the `{key}` format:

```xml
<!-- Example XML with placeholders -->
<w:t>{companyName}</w:t>
<w:t>Order Date: {orderDate}</w:t>
<w:t>Total: ${totalAmount}</w:t>
```

**Tips**:
- Place placeholders where you want dynamic content
- Keep placeholder names descriptive and meaningful
- Use camelCase for consistency (e.g., `{firstName}`, `{orderTotal}`)
- Avoid special characters in placeholder names (stick to letters and numbers)

### Relationship Management

The library automatically manages relationships for:
- Images (media files)
- Headers and footers
- Embedded content

Each relationship has:
- **ID**: Unique identifier (e.g., `rId7`)
- **Type**: Resource type (image, header, etc.)
- **Target**: Relative path to the resource

## Limitations

1. **Supported Document Elements**:
   - Text and paragraphs ✅
   - Images ✅
   - Tables ✅
   - Headers/Footers ✅
   - Charts ❌
   - SmartArt ❌
   - Embedded objects ❌

2. **Image Formats**:
   - PNG ✅
   - JPG/JPEG ✅
   - GIF ❌
   - BMP ❌
   - SVG ❌

3. **File Size**: Best suited for documents under 50MB

4. **Concurrent Operations**: Each `DocxUtil` instance should be used sequentially, not concurrently

## Troubleshooting

### Common Issues

#### 1. "File not found in assets"

**Solution**: Ensure the file is listed in `pubspec.yaml`:

```yaml
flutter:
  assets:
    - assets/docs/template.docx
```

#### 2. "Invalid image format"

**Solution**: Only PNG and JPG are supported. Convert images to these formats.

#### 3. "Placeholder not found"

**Solution**: Ensure placeholder format and names match exactly:

**Correct Format**:
- Template (Word document): `{companyName}`
- Code: `'companyName': 'ACME Corp'` (no braces in code)

**Common Mistakes**:
- ❌ Template: `{{companyName}}` (double braces)
- ❌ Template: `[companyName]` (wrong brackets)
- ❌ Code: `'{companyName}': 'ACME'` (braces in code key)
- ❌ Mismatched case: Template `{CompanyName}` vs Code `'companyName'`

**Remember**: Placeholders are case-sensitive!

#### 4. "Document corrupted after save"

**Solution**: Ensure you're not modifying the document after calling `saveDocx()`.

#### 5. Temporary files accumulating

**Solution**: Always call `dispose()` to clean up temporary files.

## Quick Reference

### Placeholder Format Cheat Sheet

| Location | Format | Example |
|----------|--------|---------|
| **Word Template** | `{key}` | `{companyName}` |
| **Dart Code (key)** | `'key'` | `'companyName'` |
| **Dart Code (value)** | `String/dynamic` | `'ACME Corp'` |

**Complete Example**:

```dart
// In your Word template: Hello {userName}, welcome to {companyName}!

await docxUtil.replaceText(
  textReplacements: {
    'userName': 'John',        // Maps to {userName}
    'companyName': 'ACME',     // Maps to {companyName}
  }
);

// Result: Hello John, welcome to ACME!
```

### Common Placeholder Types

| Type | Template Placeholder | Code Example |
|------|---------------------|--------------|
| Text | `{customerName}` | `'customerName': 'John Doe'` |
| Number | `{orderTotal}` | `'orderTotal': '1500'` |
| Date | `{invoiceDate}` | `'invoiceDate': '2025-11-07'` |
| Table | `{employeeTable}` | `'employeeTable': tableXML` |
| Image | `{companyLogo}` | `'companyLogo': MsImageUtil(...)` |

## Contributing

This library is designed to be extensible. To add new features:

1. **New Templates**: Add XML templates to `ms_word_constant.dart`
2. **New Operations**: Add methods to `DocxUtil` class
3. **New Helpers**: Create utility classes like `MsImageUtil`

## License

This project is open-source and available under the MIT License.

## Author

Anıl Kutay Uçan


**Version**: 1.0.0  
**Last Updated**: November 7, 2025

