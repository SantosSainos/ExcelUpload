package com.novellius.controller;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.ModelAttribute;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;

import com.novellius.pojo.Admin;

@Controller
public class AdminController {

	@RequestMapping("/admin")
	public String showAdmin(Model model,
			@ModelAttribute("resultado") String resultado) {

		Admin admin = new Admin();
		model.addAttribute("admin", admin);
		model.addAttribute("resultado", resultado);

		return "admin";
	}

	@RequestMapping(value = "/admin/upLoad", method = RequestMethod.POST)
	public String handleAdmin(Model model, RedirectAttributes ra,
			@RequestParam("ruta") String ruta) throws IOException {

		// Reemplaza diagonal invertida.
		ruta = ruta.replace('\\', '/');
		System.out.println("request param: " + ruta);

		// Comparación si es archivo válido o no.
		if (ruta.contains(".xlsx") || ruta.contains(".xls")) {
			// Se carga el archivo.
			FileInputStream archivo = new FileInputStream(new File(ruta));
			// Se declara el libro sobre el cual se trabajará.
			XSSFWorkbook wb = new XSSFWorkbook(archivo);
			// Se declara la hoja del libro a utilizar.
			XSSFSheet hoja = wb.getSheetAt(0);

			FormulaEvaluator formEvaluator = wb.getCreationHelper()
					.createFormulaEvaluator();

			for (Row fila : hoja) {
				for (Cell celda : fila) {
					switch (formEvaluator.evaluateInCell(celda).getCellType()) {
					case Cell.CELL_TYPE_NUMERIC:
						System.out.print(celda.getNumericCellValue() + "\t\t");
						break;
					case Cell.CELL_TYPE_STRING:
						System.out.print(celda.getStringCellValue() + "\t\t");
						break;
					}
				}
				System.out.println();
			}
			//Se cierra el libro de trabajo.
			wb.close();
			ra.addFlashAttribute("resultado", "Formato válido");
		} else {
			ra.addFlashAttribute("resultado", "No es un formato válido");
		}

		return "redirect:/admin";
	}
}
