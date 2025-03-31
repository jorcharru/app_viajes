import 'package:flutter/material.dart';
import 'package:intl/intl.dart';
import 'package:syncfusion_flutter_xlsio/xlsio.dart' as xlsio;
import 'package:path_provider/path_provider.dart';
import 'dart:io';
import 'package:open_file/open_file.dart';
import 'package:file_picker/file_picker.dart';
import 'package:share_plus/share_plus.dart';
import 'package:csv/csv.dart';
import 'dart:convert';
import 'package:shared_preferences/shared_preferences.dart';

void main() {
  runApp(AppRutas());
}

class AppRutas extends StatelessWidget {
  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      debugShowCheckedModeBanner: false,
      home: PantallaPrincipal(), // Pantalla principal por defecto
    );
  }
}

class PantallaPrincipal extends StatefulWidget {
  final double millasBase;
  final double pagoBase;
  final double pagoPorMillaExtra1;
  final double pagoPorMillaExtra2;
  final String periodoPago;
  final DateTime inicioSemana;

  PantallaPrincipal({
    this.millasBase = 16,
    this.pagoBase = 40,
    this.pagoPorMillaExtra1 = 1.76,
    this.pagoPorMillaExtra2 = 1.36,
    this.periodoPago = "semana",
    DateTime? inicioSemana,
  }) : inicioSemana = inicioSemana ?? DateTime.now();

  @override
  _PantallaPrincipalState createState() => _PantallaPrincipalState();
}

class _PantallaPrincipalState extends State<PantallaPrincipal> {
  List<Map<String, dynamic>> rutas = [];
  DateTime? fechaSeleccionada;
  TextEditingController nombreController = TextEditingController();
  TextEditingController millasController = TextEditingController();
  TextEditingController pagoAdicionalController = TextEditingController();

  String tipoRutaSeleccionada = "Programado";
  String estadoRutaSeleccionado = "Recogido";

  @override
  void initState() {
    super.initState();
    _cargarRutas();
  }

  Future<void> _cargarRutas() async {
    final prefs = await SharedPreferences.getInstance();
    final rutasJson = prefs.getStringList('rutas');
    if (rutasJson != null) {
      setState(() {
        rutas = rutasJson.map((ruta) {
          final rutaDecodificada = jsonDecode(ruta) as Map<String, dynamic>;
          // Convertir la fecha de String a DateTime
          rutaDecodificada['fecha'] = DateTime.parse(rutaDecodificada['fecha']);
          return rutaDecodificada;
        }).toList();
      });
    }
    print("Rutas cargadas: ${rutas.length}");
  }

  Future<void> _guardarRutas() async {
    final prefs = await SharedPreferences.getInstance();
    final rutasJson = rutas.map((ruta) {
      // Convertir la fecha de DateTime a String
      final rutaCodificada = Map<String, dynamic>.from(ruta);
      rutaCodificada['fecha'] = ruta['fecha'].toIso8601String();
      return jsonEncode(rutaCodificada);
    }).toList();
    await prefs.setStringList('rutas', rutasJson);
    print("Rutas guardadas: ${rutas.length}");
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(title: Text("App de Rutas")),
      body: Padding(
        padding: const EdgeInsets.all(16.0),
        child: SingleChildScrollView(
          child: Column(
            crossAxisAlignment: CrossAxisAlignment.start,
            children: [
              Text("Registrar Ruta", style: TextStyle(fontSize: 18, fontWeight: FontWeight.bold)),
              SizedBox(height: 10),
              TextField(
                controller: nombreController,
                decoration: InputDecoration(labelText: "Nombre de la Ruta"),
              ),
              TextField(
                controller: millasController,
                decoration: InputDecoration(labelText: "Millas Recorridas"),
                keyboardType: TextInputType.number,
              ),
              TextField(
                controller: pagoAdicionalController,
                decoration: InputDecoration(labelText: "Pago Adicional"),
                keyboardType: TextInputType.number,
              ),
              DropdownButtonFormField<String>(
                value: tipoRutaSeleccionada,
                items: ["Programado", "Rescate"].map((String value) {
                  return DropdownMenuItem<String>(
                    value: value,
                    child: Text(value),
                  );
                }).toList(),
                onChanged: (value) {
                  setState(() {
                    tipoRutaSeleccionada = value!;
                  });
                },
                decoration: InputDecoration(labelText: "Tipo de Ruta"),
              ),
              DropdownButtonFormField<String>(
                value: estadoRutaSeleccionado,
                items: ["Recogido", "No Recogido"].map((String value) {
                  return DropdownMenuItem<String>(
                    value: value,
                    child: Text(value),
                  );
                }).toList(),
                onChanged: (value) {
                  setState(() {
                    estadoRutaSeleccionado = value!;
                  });
                },
                decoration: InputDecoration(labelText: "Estado de Ruta"),
              ),
              SizedBox(height: 20),
              ElevatedButton(
                onPressed: () => _seleccionarFecha(context),
                child: Text(fechaSeleccionada == null ? "Seleccionar Fecha" : "Fecha: ${DateFormat('yyyy-MM-dd').format(fechaSeleccionada!)}"),
              ),
              SizedBox(height: 20),
              Center(
                child: ElevatedButton(
                  onPressed: _registrarRuta,
                  child: Text("Registrar Ruta"),
                ),
              ),
              SizedBox(height: 20),
              Text("Última Ruta Registrada", style: TextStyle(fontSize: 18, fontWeight: FontWeight.bold)),
              if (rutas.isNotEmpty)
                Container(
                  width: double.infinity,
                  padding: EdgeInsets.all(16),
                  decoration: BoxDecoration(
                    border: Border.all(color: Colors.black, width: 2),
                    borderRadius: BorderRadius.circular(8),
                  ),
                  child: Row(
                    children: [
                      Expanded(
                        child: Column(
                          crossAxisAlignment: CrossAxisAlignment.start,
                          children: [
                            Text("Nombre: ${rutas.last['nombre']}"),
                            Text("Fecha: ${DateFormat('yyyy-MM-dd').format(rutas.last['fecha'])}"),
                            Text("Millas: ${rutas.last['millas']}"),
                            Text("Pago Adicional: ${rutas.last['pagoAdicional']}"),
                          ],
                        ),
                      ),
                      Expanded(
                        child: Column(
                          crossAxisAlignment: CrossAxisAlignment.start,
                          children: [
                            Text("Pago: \$${rutas.last['pago'].toStringAsFixed(2)}"),
                            Text("Tipo: ${rutas.last['tipoRuta']}"),
                            Text("Estado: ${rutas.last['estadoRuta']}"),
                          ],
                        ),
                      ),
                    ],
                  ),
                ),
              if (rutas.isEmpty)
                Center(child: Text("No hay rutas registradas")),

              SizedBox(height: 40),

              Row(
                mainAxisAlignment: MainAxisAlignment.center,
                children: [
                  ElevatedButton(
                    onPressed: () {
                      Navigator.push(
                        context,
                        MaterialPageRoute(
                          builder: (context) => PantallaConfiguracion(
                            millasBase: widget.millasBase,
                            pagoBase: widget.pagoBase,
                            pagoPorMillaExtra1: widget.pagoPorMillaExtra1,
                            pagoPorMillaExtra2: widget.pagoPorMillaExtra2,
                            periodoPago: widget.periodoPago,
                            inicioSemana: widget.inicioSemana,
                          ),
                        ),
                      );
                    },
                    child: Text("Configuraciones"),
                  ),
                  SizedBox(width: 20),
                  ElevatedButton(
                    onPressed: () {
                      Navigator.push(
                        context,
                        MaterialPageRoute(
                          builder: (context) => PantallaRutasTablas(
                            rutas: rutas,
                            periodoPago: widget.periodoPago,
                            inicioSemana: widget.inicioSemana,
                          ),
                        ),
                      );
                    },
                    child: Text("Resumen Rutas"),
                  ),
                ],
              ),
              SizedBox(height: 10),
              Row(
                mainAxisAlignment: MainAxisAlignment.center,
                children: [
                ],
              ),
            ],
          ),
        ),
      ),
    );
  }

  Future<void> _seleccionarFecha(BuildContext context) async {
    final DateTime? seleccionada = await showDatePicker(
      context: context,
      initialDate: DateTime.now(),
      firstDate: DateTime(2023),
      lastDate: DateTime(2030),
    );
    if (seleccionada != null) {
      setState(() {
        fechaSeleccionada = seleccionada;
      });
    }
  }

  void _registrarRuta() {
    final nombre = nombreController.text;
    final millas = double.tryParse(millasController.text) ?? 0.0;
    final fecha = fechaSeleccionada ?? DateTime.now();
    final pagoAdicional = double.tryParse(pagoAdicionalController.text) ?? 0.0;

    double pago = widget.pagoBase;
    if (millas > widget.millasBase) {
      if (millas <= 25) {
        pago += (millas - widget.millasBase) * widget.pagoPorMillaExtra1;
      } else {
        pago += (25 - widget.millasBase) * widget.pagoPorMillaExtra1;
        pago += (millas - 25) * widget.pagoPorMillaExtra2;
      }
    }

    pago += pagoAdicional;

    setState(() {
      rutas.add({
        'nombre': nombre,
        'millas': millas,
        'fecha': fecha,
        'pago': pago,
        'pagoAdicional': pagoAdicional,
        'tipoRuta': tipoRutaSeleccionada,
        'estadoRuta': estadoRutaSeleccionado,
      });
    });

    _guardarRutas();

    nombreController.clear();
    millasController.clear();
    pagoAdicionalController.clear();
    fechaSeleccionada = null;
  }
}

class PantallaConfiguracion extends StatefulWidget {
  final double millasBase;
  final double pagoBase;
  final double pagoPorMillaExtra1;
  final double pagoPorMillaExtra2;
  final String periodoPago;
  final DateTime inicioSemana;

  PantallaConfiguracion({
    required this.millasBase,
    required this.pagoBase,
    required this.pagoPorMillaExtra1,
    required this.pagoPorMillaExtra2,
    required this.periodoPago,
    required this.inicioSemana,
  });

  @override
  _PantallaConfiguracionState createState() => _PantallaConfiguracionState();
}

class _PantallaConfiguracionState extends State<PantallaConfiguracion> {
  double millasBase = 16;
  double pagoBase = 40;
  double pagoPorMillaExtra1 = 1.76;
  double pagoPorMillaExtra2 = 1.36;

  String periodoPago = "semana";
  DateTime inicioSemana = DateTime.now();

  TextEditingController millasBaseController = TextEditingController();
  TextEditingController pagoBaseController = TextEditingController();
  TextEditingController pagoPorMillaExtra1Controller = TextEditingController();
  TextEditingController pagoPorMillaExtra2Controller = TextEditingController();

  String selectedDiaInicioSemana = 'Lunes';

  @override
  void initState() {
    super.initState();
    millasBaseController.text = millasBase.toString();
    pagoBaseController.text = pagoBase.toString();
    pagoPorMillaExtra1Controller.text = pagoPorMillaExtra1.toString();
    pagoPorMillaExtra2Controller.text = pagoPorMillaExtra2.toString();
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(title: Text("Configuración")),
      body: Padding(
        padding: const EdgeInsets.all(16.0),
        child: Column(
          children: [
            TextField(
              controller: millasBaseController,
              decoration: InputDecoration(labelText: "Millas Base"),
              keyboardType: TextInputType.number,
              onChanged: (value) {
                setState(() {
                  millasBase = double.tryParse(value) ?? 0.0;
                });
              },
            ),
            TextField(
              controller: pagoBaseController,
              decoration: InputDecoration(labelText: "Pago Base"),
              keyboardType: TextInputType.number,
              onChanged: (value) {
                setState(() {
                  pagoBase = double.tryParse(value) ?? 0.0;
                });
              },
            ),
            TextField(
              controller: pagoPorMillaExtra1Controller,
              decoration: InputDecoration(labelText: "Pago por Milla Extra 1"),
              keyboardType: TextInputType.number,
              onChanged: (value) {
                setState(() {
                  pagoPorMillaExtra1 = double.tryParse(value) ?? 0.0;
                });
              },
            ),
            TextField(
              controller: pagoPorMillaExtra2Controller,
              decoration: InputDecoration(labelText: "Pago por Milla Extra 2"),
              keyboardType: TextInputType.number,
              onChanged: (value) {
                setState(() {
                  pagoPorMillaExtra2 = double.tryParse(value) ?? 0.0;
                });
              },
            ),
            DropdownButton<String>(
              value: selectedDiaInicioSemana,
              items: <String>['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado', 'Domingo']
                  .map<DropdownMenuItem<String>>((String value) {
                return DropdownMenuItem<String>(
                  value: value,
                  child: Text(value),
                );
              }).toList(),
              onChanged: (String? newValue) {
                setState(() {
                  selectedDiaInicioSemana = newValue!;
                  inicioSemana = _calcularInicioSemana(newValue);
                });
              },
            ),
            ElevatedButton(
              onPressed: () {
                Navigator.pop(
                  context,
                  PantallaPrincipal(
                    millasBase: millasBase,
                    pagoBase: pagoBase,
                    pagoPorMillaExtra1: pagoPorMillaExtra1,
                    pagoPorMillaExtra2: pagoPorMillaExtra2,
                    periodoPago: periodoPago,
                    inicioSemana: inicioSemana,
                  ),
                );
              },
              child: Text("Guardar Configuración"),
            ),
          ],
        ),
      ),
    );
  }

  DateTime _calcularInicioSemana(String diaInicio) {
    DateTime now = DateTime.now();
    int diferencia = 0;

    switch (diaInicio) {
      case 'Lunes':
        diferencia = DateTime.monday - now.weekday;
        break;
      case 'Martes':
        diferencia = DateTime.tuesday - now.weekday;
        break;
      case 'Miércoles':
        diferencia = DateTime.wednesday - now.weekday;
        break;
      case 'Jueves':
        diferencia = DateTime.thursday - now.weekday;
        break;
      case 'Viernes':
        diferencia = DateTime.friday - now.weekday;
        break;
      case 'Sábado':
        diferencia = DateTime.saturday - now.weekday;
        break;
      case 'Domingo':
        diferencia = DateTime.sunday - now.weekday;
        break;
    }

    if (diferencia < 0) {
      diferencia += 7;
    }

    return now.add(Duration(days: diferencia));
  }
}

class PantallaRutasTablas extends StatefulWidget {
  final List<Map<String, dynamic>> rutas;
  final String periodoPago;
  final DateTime inicioSemana;

  PantallaRutasTablas({
    required this.rutas,
    required this.periodoPago,
    required this.inicioSemana,
  });

  @override
  _PantallaRutasTablasState createState() => _PantallaRutasTablasState();
}

class _PantallaRutasTablasState extends State<PantallaRutasTablas> {
  List<Map<String, dynamic>> rutasFiltradas = [];
  double totalGanancias = 0.0;
  String periodoSeleccionado = "semana";
  String? opcionSeleccionada;
  List<String> opcionesDisponibles = [];

  @override
  void initState() {
    super.initState();
    _actualizarOpciones();
    _aplicarFiltro();
  }

  void _actualizarOpciones() {
    setState(() {
      if (periodoSeleccionado == "semana") {
        opcionesDisponibles = _obtenerSemanasDisponibles();
      } else if (periodoSeleccionado == "mes") {
        opcionesDisponibles = _obtenerMesesDisponibles();
      } else if (periodoSeleccionado == "año") {
        opcionesDisponibles = _obtenerAnosDisponibles();
      }
      opcionSeleccionada = opcionesDisponibles.isNotEmpty ? opcionesDisponibles.first : null;
    });
  }

  List<String> _obtenerSemanasDisponibles() {
    final semanas = <String>[];
    final fechasUnicas = widget.rutas.map((ruta) => ruta['fecha']).toSet();
    for (final fecha in fechasUnicas) {
      final inicioSemana = _calcularInicioSemana(fecha);
      final finSemana = inicioSemana.add(Duration(days: 6));
      final formatoSemana = "semana ${DateFormat('dd-MM-yyyy').format(inicioSemana)} - ${DateFormat('dd-MM-yyyy').format(finSemana)}";
      if (!semanas.contains(formatoSemana)) {
        semanas.add(formatoSemana);
      }
    }
    return semanas;
  }

  List<String> _obtenerMesesDisponibles() {
    final meses = <String>[];
    final fechasUnicas = widget.rutas.map((ruta) => ruta['fecha']).toSet();
    for (final fecha in fechasUnicas) {
      final mes = DateFormat('MMMM yyyy').format(fecha);
      if (!meses.contains(mes)) {
        meses.add(mes);
      }
    }
    return meses;
  }

  List<String> _obtenerAnosDisponibles() {
    final anos = <String>[];
    final fechasUnicas = widget.rutas.map((ruta) => ruta['fecha']).toSet();
    for (final fecha in fechasUnicas) {
      final ano = DateFormat('yyyy').format(fecha);
      if (!anos.contains(ano)) {
        anos.add(ano);
      }
    }
    return anos;
  }

  DateTime _calcularInicioSemana(DateTime fecha) {
    final diaSemana = fecha.weekday;
    return fecha.subtract(Duration(days: diaSemana - DateTime.monday));
  }

  void _aplicarFiltro() {
    setState(() {
      if (opcionSeleccionada == null) return;

      rutasFiltradas = widget.rutas.where((ruta) {
        if (periodoSeleccionado == "semana") {
          final inicioSemana = _calcularInicioSemana(ruta['fecha']);
          final finSemana = inicioSemana.add(Duration(days: 6));
          final formatoSemana = "semana ${DateFormat('dd-MM-yyyy').format(inicioSemana)} - ${DateFormat('dd-MM-yyyy').format(finSemana)}";
          return formatoSemana == opcionSeleccionada;
        } else if (periodoSeleccionado == "mes") {
          final mesRuta = DateFormat('MMMM yyyy').format(ruta['fecha']);
          return mesRuta == opcionSeleccionada;
        } else if (periodoSeleccionado == "año") {
          final anoRuta = DateFormat('yyyy').format(ruta['fecha']);
          return anoRuta == opcionSeleccionada;
        }
        return false;
      }).toList();

      // Calcular el total de "Pago" (ya incluye el pago adicional)
      totalGanancias = rutasFiltradas.fold(0.0, (sum, ruta) => sum + (ruta['pago'] ?? 0.0));
    });
    print("Rutas filtradas: ${rutasFiltradas.length}");
  }

  void _eliminarRuta(int index) async {
    final confirmar = await showDialog<bool>(
      context: context,
      builder: (context) {
        return AlertDialog(
          title: Text("¿Eliminar registro?"),
          content: Text("¿Estás seguro de que deseas eliminar esta ruta?"),
          actions: [
            TextButton(
              child: Text("Cancelar"),
              onPressed: () => Navigator.pop(context, false),
            ),
            TextButton(
              child: Text("Eliminar"),
              onPressed: () => Navigator.pop(context, true),
            ),
          ],
        );
      },
    );

    if (confirmar == true) {
      final rutaEliminada = rutasFiltradas[index];

      setState(() {
        rutasFiltradas.removeAt(index); // Eliminar la fila de la lista filtrada
        totalGanancias = rutasFiltradas.fold(0.0, (sum, ruta) => sum + ruta['pago']);
      });

      // Eliminar la ruta de la lista principal
      widget.rutas.removeWhere((ruta) =>
      ruta['nombre'] == rutaEliminada['nombre'] &&
          ruta['fecha'] == rutaEliminada['fecha'] &&
          ruta['millas'] == rutaEliminada['millas'] &&
          ruta['pago'] == rutaEliminada['pago'] &&
          ruta['tipoRuta'] == rutaEliminada['tipoRuta'] &&
          ruta['estadoRuta'] == rutaEliminada['estadoRuta']);

      // Guardar los cambios en SharedPreferences
      final prefs = await SharedPreferences.getInstance();
      final rutasJson = widget.rutas.map((ruta) {
        // Convertir la fecha de DateTime a String
        final rutaCodificada = Map<String, dynamic>.from(ruta);
        rutaCodificada['fecha'] = ruta['fecha'].toIso8601String();
        return jsonEncode(rutaCodificada);
      }).toList();
      await prefs.setStringList('rutas', rutasJson);

      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(content: Text("Ruta eliminada correctamente.")),
      );
    }
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(title: Text("Resumen de Rutas")),
      body: SingleChildScrollView(
        child: Padding(
          padding: const EdgeInsets.all(16.0),
          child: Column(
            children: [
              DropdownButton<String>(
                value: periodoSeleccionado,
                items: ['semana', 'mes', 'año'].map((valor) {
                  return DropdownMenuItem<String>(
                    value: valor,
                    child: Text(valor),
                  );
                }).toList(),
                onChanged: (nuevoValor) {
                  setState(() {
                    periodoSeleccionado = nuevoValor!;
                    _actualizarOpciones();
                  });
                },
              ),
              if (opcionesDisponibles.isNotEmpty)
                DropdownButton<String>(
                  value: opcionSeleccionada,
                  items: opcionesDisponibles.map((opcion) {
                    return DropdownMenuItem<String>(
                      value: opcion,
                      child: Text(opcion),
                    );
                  }).toList(),
                  onChanged: (nuevoValor) {
                    setState(() {
                      opcionSeleccionada = nuevoValor!;
                    });
                    _aplicarFiltro();
                  },
                ),
              ElevatedButton(
                onPressed: _exportarDatos,
                child: Text("Exportar Datos"),
              ),
              SingleChildScrollView(
                scrollDirection: Axis.horizontal,
                child: DataTable(
                  border: TableBorder.all(color: Colors.black, width: 1.0),
                  columns: const [
                    DataColumn(label: Text('Nombre')),
                    DataColumn(label: Text('Fecha')),
                    DataColumn(label: Text('Millas')),
                    DataColumn(label: Text('Pago')),
                    DataColumn(label: Text('Tipo')),
                    DataColumn(label: Text('Estado')),
                    DataColumn(label: Text('Acción')),
                  ],
                  rows: rutasFiltradas.map((ruta) {
                    final index = rutasFiltradas.indexOf(ruta);
                    return DataRow(cells: [
                      DataCell(Text(ruta['nombre'] ?? '')), // Asegurar que el campo exista
                      DataCell(Text(DateFormat('yyyy-MM-dd').format(ruta['fecha']))),
                      DataCell(Text(ruta['millas']?.toString() ?? '0.0')), // Asegurar que el campo exista
                      DataCell(Text("\$${(ruta['pago'] ?? 0.0).toStringAsFixed(2)}")), // Asegurar que el campo exista
                      DataCell(Text(ruta['tipoRuta'] ?? '')), // Asegurar que el campo exista
                      DataCell(Text(ruta['estadoRuta'] ?? '')), // Asegurar que el campo exista
                      DataCell(
                        IconButton(
                          icon: Icon(Icons.delete, color: Colors.red),
                          onPressed: () => _eliminarRuta(index),
                        ),
                      ),
                    ]);
                  }).toList(),
                ),
              ),
              Text("Total Ganancias: \$${totalGanancias.toStringAsFixed(2)}"),
            ],
          ),
        ),
      ),
    );
  }

  Future<void> _exportarDatos() async {
    try {
      if (rutasFiltradas.isEmpty) {
        ScaffoldMessenger.of(context).showSnackBar(
          SnackBar(content: Text("No hay datos para exportar.")),
        );
        return;
      }

      final tipoArchivo = await showDialog<String>(
        context: context,
        builder: (context) {
          return AlertDialog(
            title: Text("Seleccione el tipo de archivo"),
            content: Column(
              mainAxisSize: MainAxisSize.min,
              children: [
                ListTile(
                  title: Text("Excel (.xlsx)"),
                  onTap: () => Navigator.pop(context, 'excel'),
                ),
                ListTile(
                  title: Text("CSV (.csv)"),
                  onTap: () => Navigator.pop(context, 'csv'),
                ),
              ],
            ),
          );
        },
      );

      if (tipoArchivo == null) return;

      final accion = await showDialog<String>(
        context: context,
        builder: (context) {
          return AlertDialog(
            title: Text("Seleccione una acción"),
            content: Column(
              mainAxisSize: MainAxisSize.min,
              children: [
                ListTile(
                  title: Text("Guardar archivo"),
                  onTap: () => Navigator.pop(context, 'guardar'),
                ),
                ListTile(
                  title: Text("Compartir archivo"),
                  onTap: () => Navigator.pop(context, 'compartir'),
                ),
              ],
            ),
          );
        },
      );

      if (accion == null) return;

      final Directory directory = await getApplicationDocumentsDirectory();
      final String path = directory.path;
      File file;

      if (tipoArchivo == 'excel') {
        final xlsio.Workbook workbook = xlsio.Workbook();
        final xlsio.Worksheet sheet = workbook.worksheets[0];

        sheet.getRangeByName('A1').setText('Nombre');
        sheet.getRangeByName('B1').setText('Fecha');
        sheet.getRangeByName('C1').setText('Millas');
        sheet.getRangeByName('D1').setText('Pago');
        sheet.getRangeByName('E1').setText('Tipo');
        sheet.getRangeByName('F1').setText('Estado');

        for (var i = 0; i < rutasFiltradas.length; i++) {
          final ruta = rutasFiltradas[i];
          sheet.getRangeByName('A${i + 2}').setText(ruta['nombre'] ?? '');
          sheet.getRangeByName('B${i + 2}').setText(DateFormat('yyyy-MM-dd').format(ruta['fecha']));
          sheet.getRangeByName('C${i + 2}').setText(ruta['millas']?.toString() ?? '0.0');
          sheet.getRangeByName('D${i + 2}').setText("\$${(ruta['pago'] ?? 0.0).toStringAsFixed(2)}");
          sheet.getRangeByName('E${i + 2}').setText(ruta['tipoRuta'] ?? '');
          sheet.getRangeByName('F${i + 2}').setText(ruta['estadoRuta'] ?? '');
        }

        final List<int> bytes = workbook.saveAsStream();
        workbook.dispose();

        file = File('$path/rutas.xlsx');
        await file.writeAsBytes(bytes);
      } else {
        final csvData = const ListToCsvConverter().convert([
          ['Nombre', 'Fecha', 'Millas', 'Pago', 'Tipo', 'Estado'], // Sin "Pago Adicional"
          ...rutasFiltradas.map((ruta) => [
            ruta['nombre'] ?? '',
            DateFormat('yyyy-MM-dd').format(ruta['fecha']),
            ruta['millas']?.toString() ?? '0.0',
            "\$${(ruta['pago'] ?? 0.0).toStringAsFixed(2)}",
            ruta['tipoRuta'] ?? '',
            ruta['estadoRuta'] ?? '',
          ]),
        ]);

        file = File('$path/rutas.csv');
        await file.writeAsString(csvData);
      }

      if (accion == 'guardar') {
        final String? savePath = await FilePicker.platform.saveFile(
          dialogTitle: 'Guardar archivo',
          fileName: tipoArchivo == 'excel' ? 'rutas.xlsx' : 'rutas.csv',
        );

        if (savePath != null) {
          await file.copy(savePath);
          ScaffoldMessenger.of(context).showSnackBar(
            SnackBar(content: Text("Archivo guardado en: $savePath")),
          );
        }
      } else {
        await Share.shareXFiles([XFile(file.path)], text: 'Aquí está el archivo exportado');
      }
    } catch (e) {
      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(content: Text("Error al exportar el archivo: $e")),
      );
    }
  }
}

