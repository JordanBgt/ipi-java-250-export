package com.example.demo.controller;

import com.example.demo.entity.Client;
import com.example.demo.entity.Facture;
import com.example.demo.entity.LigneFacture;
import com.example.demo.repository.FactureRepository;
import com.example.demo.service.ClientService;
import com.example.demo.service.FactureService;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.PrintWriter;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.Date;
import java.util.List;
import java.util.Set;

/**
 * Controlleur pour réaliser les exports.
 */
@Controller
@RequestMapping("/")
public class ExportController {

    @Autowired
    private ClientService clientService;

    @Autowired
    private FactureRepository factureRepository;


    @GetMapping("/clients/csv")
    public void clientsCSV(HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("text/csv");
        response.setHeader("Content-Disposition", "attachment; filename=\"clients.csv\"");
        PrintWriter writer = response.getWriter();
        List<Client> allClients = clientService.findAllClients();
        writer.println("Id;Nom;Prenom;Date de Naissance;Age");
        LocalDate now = LocalDate.now();
        for (Client client : allClients) {
            writer.println(
                    client.getId() + ";"
                            + client.getNom() + ";"
                            + client.getPrenom() + ";"
                            + client.getDateNaissance().format(DateTimeFormatter.ofPattern("dd/MM/yyyy")) + ";"
                            + (now.getYear() - client.getDateNaissance().getYear())
            );
        }
    }

    @GetMapping("/clients/xlsx")
    public void clientsXlsx(HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("application/vnd.ms-excel");
        response.setHeader("Content-Disposition", "attachment; filename=\"clients.xlsx\"");
        List<Client> allClients = clientService.findAllClients();
        LocalDate now = LocalDate.now();

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Clients");

        Row headerRow = sheet.createRow(0);

        Cell cellHeaderId = headerRow.createCell(0);
        cellHeaderId.setCellValue("Id");

        Cell cellHeaderPrenom = headerRow.createCell(1);
        cellHeaderPrenom.setCellValue("Prénom");

        Cell cellHeaderNom = headerRow.createCell(2);
        cellHeaderNom.setCellValue("Nom");

        Cell cellHeaderDateNaissance = headerRow.createCell(3);
        cellHeaderDateNaissance.setCellValue("Date de naissance");

        int i = 1;
        for (Client client : allClients) {
            Row row = sheet.createRow(i);

            Cell cellId = row.createCell(0);
            cellId.setCellValue(client.getId());

            Cell cellPrenom = row.createCell(1);
            cellPrenom.setCellValue(client.getPrenom());

            Cell cellNom = row.createCell(2);
            cellNom.setCellValue(client.getNom());

            Cell cellDateNaissance = row.createCell(3);
            Date dateNaissance = Date.from(client.getDateNaissance().atStartOfDay(ZoneId.systemDefault()).toInstant());
            cellDateNaissance.setCellValue(dateNaissance);

            CellStyle cellStyleDate = workbook.createCellStyle();
            CreationHelper createHelper = workbook.getCreationHelper();
            cellStyleDate.setDataFormat(createHelper.createDataFormat().getFormat("m/d/yy"));
            cellDateNaissance.setCellStyle(cellStyleDate);

            i++;
        }

        workbook.write(response.getOutputStream());
        workbook.close();

    }

    @GetMapping("/factures/xlsx")
    public void facturesXlsx(HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("application/vnd.ms-excel");
        response.setHeader("Content-Disposition", "attachment; filename=\"factures.xlsx\"");

        Workbook workbook = new XSSFWorkbook();

        List<Client> allClients = clientService.findAllClients();

        // Pour chaque client, on créer un onglet
        for (Client client : allClients) {

            Sheet clientSheet = workbook.createSheet(client.getNom());
            Row rowClientSheet = clientSheet.createRow(0);
            Cell cellPrenom = rowClientSheet.createCell(1);
            cellPrenom.setCellValue(client.getPrenom());

            Cell cellNom = rowClientSheet.createCell(2);
            cellNom.setCellValue(client.getNom());

            List<Facture> factures = factureRepository.findByClient(client);

            // Pour chaque facture, on créer un onglet
            for(Facture facture : factures){
                Sheet factureSheet = workbook.createSheet("Facture "+ facture.getId());

                int i = 1;

                // On créer un tableau pour la facture
                Set<LigneFacture> lignesFacture = facture.getLigneFactures();
                Row tableHeader = factureSheet.createRow(0);
                CellStyle tableHeaderStyle = workbook.createCellStyle();
                XSSFFont tableHeaderFont = ((XSSFWorkbook) workbook).createFont();
                tableHeaderFont.setBold(true);
                tableHeaderStyle.setBorderTop(BorderStyle.MEDIUM);
                tableHeaderStyle.setBorderRight(BorderStyle.MEDIUM);
                tableHeaderStyle.setBorderLeft(BorderStyle.MEDIUM);
                tableHeaderStyle.setBorderBottom(BorderStyle.MEDIUM);
                tableHeaderStyle.setFont(tableHeaderFont);

                Cell libelleNom = tableHeader.createCell(0);
                libelleNom.setCellValue("Nom article");
                libelleNom.setCellStyle(tableHeaderStyle);
                Cell libelleQuantite = tableHeader.createCell(1);
                libelleQuantite.setCellValue("Quantité");
                libelleQuantite.setCellStyle(tableHeaderStyle);
                Cell libellePrixUnitaire = tableHeader.createCell(2);
                libellePrixUnitaire.setCellValue("Prix unitaire");
                libellePrixUnitaire.setCellStyle(tableHeaderStyle);
                Cell libellePrixLigne = tableHeader.createCell(3);
                libellePrixLigne.setCellValue("Prix ligne");
                libellePrixLigne.setCellStyle(tableHeaderStyle);

                CellStyle tableStyle = workbook.createCellStyle();
                tableStyle.setBorderBottom(BorderStyle.THIN);
                tableStyle.setBorderTop(BorderStyle.THIN);
                tableStyle.setBorderLeft(BorderStyle.THIN);
                tableStyle.setBorderRight(BorderStyle.THIN);

                // Pour chaque ligne de facture, on remplit une ligne du tableau
                for(LigneFacture ligneFacture : lignesFacture) {

                    Row rowFactureSheet = factureSheet.createRow(i);


                    Cell cellNomArticle = rowFactureSheet.createCell(0);
                    cellNomArticle.setCellValue(ligneFacture.getArticle().getLibelle());
                    cellNomArticle.setCellStyle(tableStyle);

                    Cell cellQuantite = rowFactureSheet.createCell(1);
                    cellQuantite.setCellValue(ligneFacture.getQuantite());
                    cellQuantite.setCellStyle(tableStyle);

                    Cell cellPrixUnitaire = rowFactureSheet.createCell(2);
                    cellPrixUnitaire.setCellValue(ligneFacture.getArticle().getPrix());
                    cellPrixUnitaire.setCellStyle(tableStyle);

                    Cell cellPrixLigne = rowFactureSheet.createCell(3);
                    cellPrixLigne.setCellValue(ligneFacture.getArticle().getPrix() * ligneFacture.getQuantite());
                    cellPrixLigne.setCellStyle(tableStyle);
                    i++;
                }

                CellStyle cellTotalStyle = workbook.createCellStyle();
                cellTotalStyle.setBorderBottom(BorderStyle.MEDIUM);
                cellTotalStyle.setBottomBorderColor(IndexedColors.RED.getIndex());
                cellTotalStyle.setBorderLeft(BorderStyle.MEDIUM);
                cellTotalStyle.setLeftBorderColor(IndexedColors.RED.getIndex());
                cellTotalStyle.setBorderRight(BorderStyle.MEDIUM);
                cellTotalStyle.setRightBorderColor(IndexedColors.RED.getIndex());
                cellTotalStyle.setBorderTop(BorderStyle.MEDIUM);
                cellTotalStyle.setTopBorderColor(IndexedColors.RED.getIndex());
                XSSFFont cellTotalFont = ((XSSFWorkbook) workbook).createFont();
                cellTotalFont.setBold(true);
                cellTotalFont.setColor(IndexedColors.RED.getIndex());

                Row rowFactureSheet = factureSheet.createRow(i);
                Cell cellTotalLibelle = rowFactureSheet.createCell(0);
                cellTotalLibelle.setCellValue("Total : ");
                cellTotalLibelle.setCellStyle(cellTotalStyle);

                Cell cellTotal = rowFactureSheet.createCell(1);
                cellTotal.setCellValue(facture.getTotal());
                cellTotal.setCellStyle(cellTotalStyle);
                cellTotalStyle.setFont(cellTotalFont);
            }
        }

        workbook.write(response.getOutputStream());
        workbook.close();
    }
}
