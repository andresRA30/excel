import 'dart:io';

import 'package:flutter/material.dart';
import 'package:syncfusion_flutter_xlsio/xlsio.dart';
import 'package:path_provider/path_provider.dart';
import 'package:open_file/open_file.dart';
import 'package:universal_html/html.dart' show AnchorElement;
import 'package:flutter/foundation.dart' show kIsWeb;
import 'dart:convert';
import 'package:spreadsheet_decoder/spreadsheet_decoder.dart';
import 'package:file_picker/file_picker.dart';

void main() {
  runApp(const MyApp());
}

class MyApp extends StatelessWidget {
  const MyApp({Key? key}) : super(key: key);

  // This widget is the root of your application.
  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'Flutter  ',
      theme: ThemeData(
        primarySwatch: Colors.blue,
      ),
      home: const MyHomePage(title: 'Flutter   Page'),
    );
  }
}

class MyHomePage extends StatefulWidget {
  const MyHomePage({Key? key, required this.title}) : super(key: key);

  final String title;

  @override
  State<MyHomePage> createState() => _MyHomePageState();
}

class _MyHomePageState extends State<MyHomePage> {
  @override
  Widget build(BuildContext context) {
    return Scaffold(
      body: Center(
        child: ElevatedButton(
          child: Text('Import Excel'),
          // onPressed: exportExcel,
          onPressed: importExcel,
        ),
      ),
    );
  }

  Future<void> exportExcel() async {
// Create a new Excel Document.
    final Workbook workbook = Workbook();

// Accessing worksheet via index.
    final Worksheet sheet = workbook.worksheets[0];

// Set the text value Titles.
    sheet.getRangeByName('A1').setText('Animales');
    sheet.getRangeByName('B1:C1').setText('Edad promedio');
    sheet.getRangeByName('C1').setText('Peso total');
    sheet.getRangeByName('D1').setText('Peso promedio');
    sheet.getRangeByName('E1').setText('Forraje diario requerido');
    sheet.getRangeByName('F1').setText('Forraje requerido promedio');
//Initialize the List\<Object>

    Map estructura = {
      'animales': 0.0,
      'edad': 0,
      'pesoTotal': 0,
      'pesoPronedio': 0,
      'forrajeDiario': 0,
      'forrajeRequerido': 0,
    };
    sheet.getRangeByName('A2').setText(estructura['animales'].toString());
    sheet.getRangeByName('B2').setText(estructura['edad'].toString());
    sheet.getRangeByName('C2').setText(estructura['pesoTotal'].toString());
    sheet.getRangeByName('D2').setText(estructura['pesoPronedio'].toString());
    sheet.getRangeByName('E2').setText(estructura['forrajeDiario'].toString());
    sheet
        .getRangeByName('F2')
        .setText(estructura['forrajeRequerido'].toString());
// Save and dispose the document.
    final List<int> bytes = workbook.saveAsStream();
    workbook.dispose();

// Get external storage directory
    final directory = await getApplicationDocumentsDirectory();

// Get directory path
    final path = directory.path;
    print(path);
// Create an empty file to write Excel data
    File file = File('$path/Report.xlsx');

// Write Excel data
    await file.writeAsBytes(bytes, flush: true);

// Open the Excel document in mobile
    OpenFile.open('$path/Report.xlsx');
  }

  Future<void> importExcel() async {
    FilePickerResult? result = await FilePicker.platform.pickFiles();
    print('${result?.files.single.path}');
    if (result != null) {
      // File file = File(result.files.single.path);
    } else {
      // User canceled the picker
    }
    var bytes = File('${result?.files.single.path}').readAsBytesSync();
    var excel = SpreadsheetDecoder.decodeBytes(bytes);
    var data = [];

    for (var table in excel.tables.keys) {
      // print(table); //sheet Name
      // print(excel.tables[table]!.maxCols);
      // print(excel.tables[table]!.maxRows);
      for (var row in excel.tables[table]!.rows) {
        data.add(row);
        print("$row");
      }
    }
    print("hola");
    print(data[0][0]);
  }
}
