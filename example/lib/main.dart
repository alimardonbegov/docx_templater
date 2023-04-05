import 'dart:io';
import 'package:http/http.dart';
import 'package:tpl_docx/tpl_docx.dart';

void main(List<String> args) async {
  Future<List<int>> getDocBytesFromRemote(String docxUrl) async {
    try {
      final Client _client = Client();
      final Response req = await _client.get(Uri.parse(docxUrl));
      final List<int> bytes = req.bodyBytes;
      return bytes;
    } catch (e) {
      throw Exception('error downloading remote .docx template file: ' + e.toString());
    }
  }

  Future<List<int>> getDocBytesFromLocal(String docPath) async {
    try {
      final file = File(docPath);
      final List<int> bytes = file.readAsBytesSync();
      return bytes;
    } catch (e) {
      throw Exception('error downloading local .docx template file: ' + e.toString());
    }
  }

  final Map<String, String> testMap = {
    "DOC_NUMBER": "22222",
    "IME_PREZIME": "BBBBB",
  };

  final Map<String, String> testMapIzjava = {
    "COMPANY_NAME": "World world N",
    "COMPANY_PIB": "222299999",
  };

  final List<int> bytes = await getDocBytesFromLocal("lib/templates/izjava.docx");
  final tpl = Templater(bytes);

  tpl.writeMergedFields(testMapIzjava);
  final List<String> mergedFields = tpl.mergedFields;
  print(mergedFields);
  final List<int> generatedBytes = tpl.getGeneratedBytes()!;

  final file = File("lib/results/result.docx");
  file.writeAsBytesSync(generatedBytes);
}
