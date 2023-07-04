import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import javax.swing.JOptionPane;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Reembolsos {

    public static void main(String[] args) {
        try {
            new Reembolsos();
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    }

    public Reembolsos() throws IOException {
       // Ajuste o caminho das pastas nas duas linhas abaixo, para o que vc está usando localmente no seu computador
        String pastaSolicitacoesReembolso = "C:/Solicitações de Reembolso";
        String pastaConsolidadoReembolso = "C:/Consolidado de Reembolso";
        String nomeSolicitante = "";
        double totalSolicitante = 0;
        double totalGeral = 0;
        int linhaSaida = 4;

        // Vamos ler todos os arquivos da pasta de solicitações de reembolso

        File arquivo[];
        File todas = new File(pastaSolicitacoesReembolso);
        arquivo = todas.listFiles();

        if (arquivo.length == 0) {
      // Ajuste o caminho da pasta que vc está usando localmente no seu computador      
            JOptionPane.showMessageDialog(null, "Não há planilhas de reembolso na pasta C:/Solicitações de Reembolso. Não haverá atualização da planilha de consolidado de reembolsos.");

        } else {

            for(int i = 0; i < arquivo.length; i++){

                   // Abre planilha de solicitações de reembolsos    

                   FileInputStream fis = null;
                   try {
                        fis = new FileInputStream(arquivo[i]);
                   } catch (FileNotFoundException e) {
                        // TODO Auto-generated catch block
                        e.printStackTrace();
                   }
                   XSSFWorkbook workbookEntrada=new XSSFWorkbook(fis);
                   XSSFSheet sheetEntrada=workbookEntrada.getSheetAt(0);

                   // Lê nome do solicitante e total de solicitação de reembolso do mesmo

                   nomeSolicitante = sheetEntrada.getRow(3).getCell(2).getStringCellValue();
                   totalSolicitante = sheetEntrada.getRow(16).getCell(3).getNumericCellValue();

                   // Acumula valor de reembolso solicitado

                   totalGeral = totalGeral + totalSolicitante;

                   // Grava linha na planilha de consolidado de reembolsos

                   try {
                          String filePath = pastaConsolidadoReembolso + "/" + "Consolidado de reembolso de despesas.xlsx";
                       File file=new File(filePath);
                       FileInputStream arq=new FileInputStream(file);
                       XSSFWorkbook workbookSaida=new XSSFWorkbook(arq);
                       XSSFSheet sheetSaida=workbookSaida.getSheetAt(0);

                       if(linhaSaida==4) {
                          // Limpeza das linhas do processamento anterior na planilha de consolidado de reembolsos
                             for(int j = 4; j <= 50; j++){
                               sheetSaida.getRow(j).createCell(1).setCellValue("");
                               sheetSaida.getRow(j).createCell(2).setCellValue("");
                             }
                       }

                       // Grava nome do solicitante e total de solicitação de reembolso do mesmo

                       sheetSaida.getRow(linhaSaida).createCell(1).setCellValue(nomeSolicitante);  
                       NumberFormat formataValor = new DecimalFormat(",###.00");     
                       sheetSaida.getRow(linhaSaida).createCell(2).setCellValue(formataValor.format(totalSolicitante)); 

                       linhaSaida++;

                       // Gravação de linha de total geral

                       if (i==arquivo.length-1) {
                           sheetSaida.getRow(linhaSaida).createCell(1).setCellValue("TOTAL GERAL = ");
                           sheetSaida.getRow(linhaSaida).createCell(2).setCellValue(formataValor.format(totalGeral));

                           // Formatação da literal de TOTAL GERAL à direita

                           final XSSFCellStyle style = workbookSaida.createCellStyle();
                           style.setAlignment(HorizontalAlignment.RIGHT);
                           XSSFRow row = sheetSaida.getRow(linhaSaida);
                           XSSFCell cell = row.getCell((short) 1);
                           cell.setCellStyle(style);
                       }

                       FileOutputStream fos=new FileOutputStream(filePath);
                       workbookSaida.write(fos);
                       fos.close();
                       workbookSaida.close();

                   } 
                      catch (FileNotFoundException saida) {
                         saida.printStackTrace();
                   }

                   // Fechamento da planilha de reembolso lida
                   workbookEntrada.close();

                } // Fechamento do FOR de leitura de planilhas de reembolso
        }

    }

}
