import 'dart:convert';
import 'dart:io';
import 'package:archive/archive_io.dart';
import 'package:path/path.dart' as path;

class Templater {
  final List<int> bytes;
  final Map<String, String> map;

  List<String> _mergedFields = [];
  late Archive _zip;
  late String _docXml;
  late int _documentXmlIndex;
  late Directory _dir;

  Templater({
    required this.bytes,
    required this.map,
  });

  void _templateParse(String text) {
    final List<String> fields = [];
    final RegExp re = RegExp('{{\\w*}}', caseSensitive: true, multiLine: true);
    final Iterable<Match> matches = re.allMatches(text);

    if (matches.isEmpty) {
      fields;
    } else {
      for (var match in matches) {
        final int group = match.groupCount;
        final String field = match.group(group)!;
        final String firstChunk = field.replaceAll('{{', '').trim(); // remove templating braces
        final String secChunk = firstChunk.replaceAll('}}', '').trim(); // remove templating braces
        fields.add(secChunk);
      }
    }
    _mergedFields = fields.toSet().toList();
  }

  void _getArchiveAndXmlString() {
    _zip = ZipDecoder().decodeBytes(bytes, verify: true);

    _documentXmlIndex = _zip.files.indexWhere((file) => file.name == 'word/document.xml');

    final ArchiveFile documentXml = _zip.files[_documentXmlIndex];
    final List<int> content = documentXml.content as List<int>;

    _docXml = utf8.decode(content);
  }

  void _writeMergeFields() {
    for (var field in _mergedFields) {
      if (map.containsKey(field) &&
          _docXml.contains(
            RegExp(
              '{{\\w*}}',
              caseSensitive: true,
              multiLine: true,
            ),
          )) {
        _docXml = _docXml.replaceAll(RegExp('{{$field}}'), map[field]!);
      }
    }
  }

  List<String> get mergeFields => _mergedFields; //get template fields

  List<int>? generateTpl() {
    _getArchiveAndXmlString();
    _templateParse(_docXml);

    _writeMergeFields();

    final Archive newZip = Archive();
    for (var file in _zip.files) {
      if (file.name != 'word/document.xml') {
        newZip.addFile(file);
      } else {
        final List<int> xmlBytes = utf8.encode(_docXml);
        final ArchiveFile newWordDocumentXml = ArchiveFile('word/document.xml', xmlBytes.length, xmlBytes);
        newZip.addFile(newWordDocumentXml);
      }
    }

    final newBytes = ZipEncoder().encode(newZip);
    return newBytes;
  }
}
