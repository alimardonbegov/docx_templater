import 'dart:io';
import 'dart:convert';
import 'package:collection/collection.dart' show IterableExtension;
import 'package:archive/archive_io.dart';
import 'package:http/http.dart';
import 'package:xml/xml.dart';
import './docx_constants.dart';
import 'tpl_response.dart';

// import 'package:flutter/services.dart' show rootBundle;//! package for local download
// import 'package:path/path.dart' as path;//! package for local download

class DocxTpl {
  final String? docxTemplate; //file path or url to the .docx file
  final bool isRemoteFile;
  final bool isAssetFile;
  late Archive _zip; // internal zip file object of the read .docx file
  late List<int> _bytes; // hold doctemplate bytes (//!created new )
  late List<int> _xmlBytes; //!created new

  Map<dynamic, XmlDocument> _parts = Map<dynamic, XmlDocument>(); // hold xml document
  List<XmlElement> _instrTextChildren = [];
  List<String> mergedFields = []; // hold merge fields extracted from the document
  var _settings;
  var _settingsInfo;

  DocxTpl({
    this.docxTemplate,
    this.isAssetFile: false,
    this.isRemoteFile: false,
  });

// Future saveAssetTpl(String assetPath) async {
//   try {
//     final assetTpl = await rootBundle.load(assetPath);
//     final List<int> bytes = assetTpl.buffer.asUint8List();
//     return bytes;
//   } catch (e) {
//     return e.toString();
//   }
// }

// без использования файловой системы
  Future<List<int>> _docxRemoteFileDownloader(String docxUrl) async {
    try {
      final Client _client = Client();
      final Response req = await _client.get(Uri.parse(docxUrl));
      final List<int> bytes = req.bodyBytes;
      return bytes;
    } catch (e) {
      throw Exception('error downloading remote .docx template file: ' + e.toString());
    }
  }

// без использования файловой системы
  List<String> _templateParse(String text) {
    final List<String> fields = [];
    final RegExp re = RegExp('{{\\w*}}', caseSensitive: true, multiLine: true);
    final Iterable<Match> matches = re.allMatches(text);

    if (matches.isEmpty) {
      return fields;
    } else {
      for (var match in matches) {
        final int group = match.groupCount;
        final String field = match.group(group)!;
        final String firstChunk = field.replaceAll('{{', '').trim(); // remove templating braces
        final String secChunk = firstChunk.replaceAll('}}', '').trim(); // remove templating braces
        fields.add(secChunk);
      }
    }
    return fields;
  }

// без использования файловой системы
  List _getTreeOfFile(XmlElement file) {
    final String type = file.getAttribute('PartName')!;
    final String innerFile = type.replaceFirst('/', '');
    final ArchiveFile zi = _zip.findFile(innerFile)!;
    final List<int> ziFileData = zi.content as List<int>;
    final String text = utf8.decode(ziFileData);
    final XmlDocument parsedZi = XmlDocument.parse(text);

    return [zi, parsedZi];
  }

// без использования файловой системы
  Future<List<int>> save() async {
    try {
      ArchiveFile xmlFile = ArchiveFile('generated_tpl.xml', _xmlBytes.length, _xmlBytes);
      _zip.clear();
      _zip.addFile(xmlFile);
      return ZipEncoder().encode(_zip)!;
    } catch (e) {
      throw Exception('failed to save generated file: ${e.toString()}');
    }
  }

// без использования файловой системы
  List<String> getMergeFields() {
    return mergedFields.toSet().toList();
  }

// без использования файловой системы
  Future<void> writeMergeFields({required Map<String, dynamic> data}) async {
    final List<String> fields = getMergeFields();
    final List<XmlElement> elementTags = _instrTextChildren.toSet().toList(); //remove any duplicates if any

    for (var field in fields) {
      // replace field with proper data in elText
      for (var element in elementTags) {
        // grab the text to check templating {{..}} and change field
        var elText = element.text;

        // only change proper templated fields  {{..}} and leave the rest as is
        if (elText.contains(RegExp(
          '{{\\w*}}',
          caseSensitive: true,
          multiLine: true,
        ))) {
          final String rep = elText.replaceAll(RegExp('{{$field}}'), data[field]);
          element.innerText = rep;
        }
      }
    }

    // grab any element's root note and save to disk
    // grab the root document already changed by calling [element.innerText] = '<new-data>' while replacing fields above
    final XmlDocument documentXmlRootDoc = elementTags.first.root.root.document!;
    final String xmlString = documentXmlRootDoc.toXmlString(); // grab xml as is without pretty printed
    _xmlBytes = utf8.encode(xmlString);

    // final String docXml = path.join(_dir.path, 'word', 'document.xml'); // write document to temp dir file
    // final File docXmlFile = await File(docXml).create(recursive: true);
    // await docXmlFile.writeAsString(xml);
  }

// без использования файловой системы
  Future<MergeResponse> parseDocxTpl() async {
    try {
      if (isRemoteFile) _bytes = await _docxRemoteFileDownloader(docxTemplate!);
      // if (!isAssetFile && !isRemoteFile) _docxFile = File(docxTemplate!);
      // if (isAssetFile) _bytes = await _saveAssetTpl(docxTemplate!);

      _zip = ZipDecoder().decodeBytes(_bytes);

      ArchiveFile? zippedFile =
          _zip.files.firstWhereOrNull((zippedElement) => zippedElement.name == '[Content_Types].xml');

      if (zippedFile == null) {
        throw Exception('failed to read .docx template file passed');
      }

      if (zippedFile.isFile) {
        final List<int> fileData = zippedFile.content as List<int>;
        final String text = Utf8Decoder().convert(fileData);
        final contentTypes = XmlDocument.parse(text);

        // loop through xml document to check required data
        for (var file in contentTypes.findAllElements('Override', namespace: "${NAMESPACES['ct']}")) {
          var type = file.getAttribute('ContentType', namespace: "${NAMESPACES['ct']}");

          for (var contentTypePart in CONTENT_TYPES_PARTS) {
            if (type == contentTypePart) {
              // checking
              var chunkResp = _getTreeOfFile(file);
              _parts[chunkResp.first] = chunkResp.last;
            }
          }

          // check in another
          if (type == CONTENT_TYPE_SETTINGS) {
            var chunkResp = _getTreeOfFile(file);
            _settingsInfo = chunkResp.first;
            _settings = chunkResp.last;
          }

          for (var part in _parts.values) {
            // hunt for w:t text and check for simple templating {{<name>}}
            for (var parent in part.findAllElements('w:t')) {
              _instrTextChildren.add(parent);
            }

            _instrTextChildren.toSet().toList(); // use unique fields

            for (var instrChild in _instrTextChildren) {
              List<String> chunkResult = _templateParse(instrChild.text); // extract merge-field
              mergedFields.addAll(chunkResult); // add merge fields to list
            }
          }
        }
      }

      return MergeResponse(
        mergeStatus: MergeResponseStatus.Success,
        message: 'success',
      );
    } catch (e) {
      print(e.toString());
      return MergeResponse(
        mergeStatus: MergeResponseStatus.Error,
        message: e.toString(),
      );
    }
  }
}
