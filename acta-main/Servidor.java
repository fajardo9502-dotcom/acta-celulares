import com.sun.net.httpserver.HttpServer;
import com.sun.net.httpserver.HttpHandler;
import com.sun.net.httpserver.HttpExchange;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.ss.usermodel.CellType;

import java.io.*;
import java.net.InetSocketAddress;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.Base64;
import java.util.LinkedHashMap;
import java.util.Map;

public class Servidor {

    private static final String EXCEL_PATH = "C:\\Users\\1030650138\\OneDrive - Colombiana de Comercio S.A\\base actas de entrega.xlsx";
    private static final String PDF_FOLDER = "PDFs";

    public static void main(String[] args) throws IOException {
        new File(PDF_FOLDER).mkdirs();

        HttpServer server = HttpServer.create(new InetSocketAddress(8080), 0);

        server.createContext("/", new ArchivoEstatico());
        server.createContext("/api/acta", new RecibirActa());

        server.setExecutor(null);
        System.out.println("----------------------------------------------");
        System.out.println("SERVIDOR LISTO EN: http://localhost:8080");
        System.out.println("----------------------------------------------");
        server.start();
    }

    /** Sirve archivos estaticos (HTML, CSS, JS, imagenes). */
    static class ArchivoEstatico implements HttpHandler {
        @Override
        public void handle(HttpExchange exchange) throws IOException {
            String ruta = exchange.getRequestURI().getPath();
            if (ruta.equals("/")) ruta = "/index.html";

            Path base = new File(".").getCanonicalFile().toPath();
            Path archivo = new File("." + ruta).getCanonicalFile().toPath();

            if (!archivo.startsWith(base)) {
                exchange.sendResponseHeaders(403, -1);
                return;
            }

            if (Files.exists(archivo) && !Files.isDirectory(archivo)) {
                byte[] contenido = Files.readAllBytes(archivo);
                exchange.getResponseHeaders().set("Content-Type", detectarTipo(ruta));
                exchange.sendResponseHeaders(200, contenido.length);
                try (OutputStream os = exchange.getResponseBody()) {
                    os.write(contenido);
                }
            } else {
                byte[] msg = "404 - Archivo no encontrado".getBytes(StandardCharsets.UTF_8);
                exchange.sendResponseHeaders(404, msg.length);
                try (OutputStream os = exchange.getResponseBody()) {
                    os.write(msg);
                }
            }
        }
    }

    /** Recibe datos del formulario y PDF, guarda en Excel y carpeta PDFs. */
    static class RecibirActa implements HttpHandler {
        @Override
        public void handle(HttpExchange exchange) throws IOException {
            exchange.getResponseHeaders().add("Access-Control-Allow-Origin", "*");
            exchange.getResponseHeaders().add("Access-Control-Allow-Methods", "POST, OPTIONS");
            exchange.getResponseHeaders().add("Access-Control-Allow-Headers", "Content-Type");

            if ("OPTIONS".equalsIgnoreCase(exchange.getRequestMethod())) {
                exchange.sendResponseHeaders(204, -1);
                return;
            }

            if (!"POST".equalsIgnoreCase(exchange.getRequestMethod())) {
                exchange.sendResponseHeaders(405, -1);
                return;
            }

            try {
                String body = new String(exchange.getRequestBody().readAllBytes(), StandardCharsets.UTF_8);
                Map<String, String> datos = parsearJson(body);

                System.out.println("\n--- DATOS RECIBIDOS ---");
                datos.forEach((k, v) -> {
                    if (!"pdfBase64".equals(k) && !"firma_digital".equals(k)) {
                        System.out.println("  " + k + " = " + v);
                    }
                });

                // 1. Guardar PDF
                String nombrePdf = generarNombrePdf(datos);
                String pdfBase64 = datos.getOrDefault("pdfBase64", "");
                String rutaPdf = "";

                if (!pdfBase64.isEmpty()) {
                    if (pdfBase64.contains(",")) {
                        pdfBase64 = pdfBase64.substring(pdfBase64.indexOf(",") + 1);
                    }
                    byte[] pdfBytes = Base64.getDecoder().decode(pdfBase64);
                    File archivoPdf = new File(PDF_FOLDER, nombrePdf);
                    try (FileOutputStream fos = new FileOutputStream(archivoPdf)) {
                        fos.write(pdfBytes);
                    }
                    rutaPdf = archivoPdf.getPath();
                    System.out.println("  PDF guardado: " + rutaPdf);
                }

                // 2. Agregar fila al Excel
                agregarFilaExcel(datos, rutaPdf);
                System.out.println("  Excel actualizado: " + EXCEL_PATH);
                System.out.println("--------------------------------------\n");

                String response = "OK: PDF guardado y Excel actualizado";
                byte[] responseBytes = response.getBytes(StandardCharsets.UTF_8);
                exchange.getResponseHeaders().set("Content-Type", "text/plain; charset=UTF-8");
                exchange.sendResponseHeaders(200, responseBytes.length);
                try (OutputStream os = exchange.getResponseBody()) {
                    os.write(responseBytes);
                }

            } catch (Exception e) {
                System.err.println("Error procesando acta: " + e.getMessage());
                e.printStackTrace();
                String error = "Error: " + e.getMessage();
                byte[] errorBytes = error.getBytes(StandardCharsets.UTF_8);
                exchange.sendResponseHeaders(500, errorBytes.length);
                try (OutputStream os = exchange.getResponseBody()) {
                    os.write(errorBytes);
                }
            }
        }
    }

    /**
     * Agrega una fila al Libro1.xlsx con los datos del formulario.
     * Columnas del Excel (fila 1 = encabezados, fila 2+ = datos):
     * A:IMEI1, B:IMEI2, C:MODELO, D:C.Costos, E:Fecha Asignacion,
     * F:UNI, G:UN, H:Supervisor, I:Zona Cargo, J:Codigo, K:Cedula,
     * L:Funcionario, M:Bateria, N:Cargador, O:TIPO, P:Ganga Linea,
     * Q:Tipo De Novedades, R:Estado, S:Tipo Plan, T:Costo Plan,
     * U:(vacio), V:Nombre Cuenta, W:Tipo Cargo, X:Tipo Legis,
     * Y:ANTERIOR USUARIO, Z:Ruta PDF
     */
    static synchronized void agregarFilaExcel(Map<String, String> datos, String rutaPdf) throws IOException {
        File excelFile = new File(EXCEL_PATH);
        XSSFWorkbook workbook;

        // Leer el archivo si existe, o crear uno nuevo
        if (excelFile.exists()) {
            try (FileInputStream fis = new FileInputStream(excelFile)) {
                workbook = new XSSFWorkbook(fis);
            }
        } else {
            workbook = new XSSFWorkbook();
            workbook.createSheet("base"); // Crea la hoja si es nuevo
        }

        XSSFSheet sheet = workbook.getSheetAt(0);
        
        // --- LOGICA DE AUTOINCREMENTO Y FILA ---
        int lastRowIndex = sheet.getLastRowNum(); 
        // Si el archivo está vacío o solo tiene el encabezado, getLastRowNum suele ser 0
        XSSFRow row = sheet.createRow(lastRowIndex + 1);
        
        // El número de No (Columna A) será el índice de la fila actual (lastRowIndex + 1)
        int numeroActa = lastRowIndex + 1;

        String fechaHoy = LocalDate.now().format(DateTimeFormatter.ofPattern("dd/MM/yyyy"));
        // Columna | indice | Campo
        row.createCell(0).setCellValue(numeroActa);                          // A - No (AUTOINCREMENTO)
        row.createCell(1).setCellValue(datos.getOrDefault("Telefono", ""));   // B - Telefono
        row.createCell(2).setCellValue(datos.getOrDefault("IMEI1", ""));      // C - IMEI1
        row.createCell(3).setCellValue(datos.getOrDefault("IMEI2", ""));      // D - IMEI2
        row.createCell(4).setCellValue(datos.getOrDefault("MODELO", ""));     // E - MODELO
        row.createCell(5).setCellValue(datos.getOrDefault("C. Costos", ""));  // F - C. Costos
        row.createCell(6).setCellValue(fechaHoy);                             // G - Fecha Asignacion
        
        row.createCell(7).setCellValue("");                                   // H - UN2
        row.createCell(8).setCellValue(datos.getOrDefault("Zona o Cargo", ""));// I - UN (Ej: pyg547)
        row.createCell(9).setCellValue(datos.getOrDefault("Supervisor", ""));  // J - Supervisor
        row.createCell(10).setCellValue(datos.getOrDefault("Zona o Cargo", ""));// K - Zona o Cargo
        row.createCell(11).setCellValue(datos.getOrDefault("Código", ""));     // L - Código
        row.createCell(12).setCellValue(datos.getOrDefault("Cedula", ""));     // M - Cedula
        row.createCell(13).setCellValue(datos.getOrDefault("Funcionario", ""));// N - Funcionario
        row.createCell(14).setCellValue(datos.getOrDefault("Bateria", "No"));  // O - Bateria
        row.createCell(15).setCellValue(datos.getOrDefault("Cargador", ""));
        row.createCell(16).setCellValue(datos.getOrDefault("TipoEquipo", ""));
        row.createCell(17).setCellValue("Activo");
        row.createCell(18).setCellValue(datos.getOrDefault("Novedades", ""));
        row.createCell(19).setCellValue("Activo");
        row.createCell(20).setCellValue(""); 
        row.createCell(21).setCellValue(datos.getOrDefault("Tipo Plan", ""));
        row.createCell(22).setCellValue(datos.getOrDefault("Costo Plan", ""));
        row.createCell(23).setCellValue(datos.getOrDefault("Cuenta", ""));
        row.createCell(24).setCellValue(datos.getOrDefault("Nombre Cuenta", ""));
        row.createCell(25).setCellValue(datos.getOrDefault("Tipo Cargo", ""));
        row.createCell(26).setCellValue(datos.getOrDefault("Tipo Logia", ""));
        row.createCell(27).setCellValue("");
        row.createCell(28).setCellValue(rutaPdf);

        // Guardar cambios
        try (FileOutputStream fos = new FileOutputStream(excelFile)) {
            workbook.write(fos);
        } finally {
            workbook.close();
        }
    }

    /** Genera nombre unico para el PDF: acta_CEDULA_YYYYMMDD.pdf */
    static String generarNombrePdf(Map<String, String> datos) {
        String cedula = datos.getOrDefault("Cedula", "sin_cedula").replaceAll("[^a-zA-Z0-9]", "");
        String fecha = LocalDate.now().format(DateTimeFormatter.ofPattern("yyyyMMdd"));
        return "acta_" + cedula + "_" + fecha + ".pdf";
    }

    /** Parser JSON simple para objetos planos {"key":"value",...}. */
    static Map<String, String> parsearJson(String json) {
        Map<String, String> mapa = new LinkedHashMap<>();
        json = json.trim();
        if (json.startsWith("{")) json = json.substring(1);
        if (json.endsWith("}")) json = json.substring(0, json.length() - 1);

        int i = 0;
        while (i < json.length()) {
            // Buscar clave
            int comillaInicio = json.indexOf('"', i);
            if (comillaInicio == -1) break;
            int comillaFin = json.indexOf('"', comillaInicio + 1);
            if (comillaFin == -1) break;
            String clave = json.substring(comillaInicio + 1, comillaFin);

            // Buscar dos puntos
            int dosPuntos = json.indexOf(':', comillaFin);
            if (dosPuntos == -1) break;

            // Buscar valor
            int valorInicio = dosPuntos + 1;
            while (valorInicio < json.length() && json.charAt(valorInicio) == ' ') valorInicio++;

            String valor;
            if (valorInicio < json.length() && json.charAt(valorInicio) == '"') {
                // Valor string - buscar cierre respetando escapes
                StringBuilder sb = new StringBuilder();
                int j = valorInicio + 1;
                while (j < json.length()) {
                    char c = json.charAt(j);
                    if (c == '\\' && j + 1 < json.length()) {
                        char next = json.charAt(j + 1);
                        if (next == '"') { sb.append('"'); j += 2; continue; }
                        if (next == '\\') { sb.append('\\'); j += 2; continue; }
                        if (next == 'n') { sb.append('\n'); j += 2; continue; }
                        if (next == '/') { sb.append('/'); j += 2; continue; }
                    }
                    if (c == '"') break;
                    sb.append(c);
                    j++;
                }
                valor = sb.toString();
                i = j + 1;
            } else {
                // Valor no-string (null, number, boolean)
                int finValor = json.indexOf(',', valorInicio);
                if (finValor == -1) finValor = json.length();
                valor = json.substring(valorInicio, finValor).trim();
                if ("null".equals(valor)) valor = "";
                i = finValor;
            }

            mapa.put(clave, valor);

            // Avanzar pasando la coma
            int coma = json.indexOf(',', i);
            if (coma == -1) break;
            i = coma + 1;
        }
        return mapa;
    }

    /** Detecta Content-Type segun extension. */
    static String detectarTipo(String ruta) {
        if (ruta.endsWith(".html")) return "text/html; charset=UTF-8";
        if (ruta.endsWith(".css"))  return "text/css; charset=UTF-8";
        if (ruta.endsWith(".js"))   return "application/javascript; charset=UTF-8";
        if (ruta.endsWith(".png"))  return "image/png";
        if (ruta.endsWith(".jpg") || ruta.endsWith(".jpeg")) return "image/jpeg";
        if (ruta.endsWith(".ico"))  return "image/x-icon";
        if (ruta.endsWith(".pdf"))  return "application/pdf";
        return "application/octet-stream";
    }
}
