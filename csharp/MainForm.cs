/*
 * Создано в SharpDevelop.
 * Пользователь: egor
 * Дата: 30.08.2017
 * Время: 18:10
 * 
 * Для изменения этого шаблона используйте меню "Инструменты | Параметры | Кодирование | Стандартные заголовки".
 */
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;

namespace badge
{
	/// <summary>
	/// Description of MainForm.
	/// </summary>
	public partial class MainForm : Form
	{
		private string wordFile;
		private string excelFile;
		private List<Student> fios;
		
		public MainForm()
		{
			//
			// The InitializeComponent() call is required for Windows Forms designer support.
			//
			InitializeComponent();
			
			this.wordFile = "";
			this.excelFile = "";
			this.fios = new List<Student>();
		}
		
		void Button2Click(object sender, EventArgs e)
		{
			// Create an instance of the open file dialog box.
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Title = "Выберите шаблон Word для заполнения бейджей";
            openFileDialog1.Filter = "Файлы Word (.docx) |*.docx| Файлы Word 97-2003 (.doc) |*.doc";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.Multiselect = false;

            // Call the ShowDialog method to show the dialog box.
            DialogResult userClickedOK = openFileDialog1.ShowDialog();
			
            // Process input if the user clicked OK.
            if (userClickedOK == DialogResult.OK)
            {
            	this.wordFile = openFileDialog1.FileName;
            	this.textBox2.Text = openFileDialog1.SafeFileName;
            }	
	
		}
		void Button1Click(object sender, EventArgs e)
		{
			// Create an instance of the open file dialog box.
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Title = "Выберите файл Excel с данными для заполнения бейджей";
            openFileDialog1.Filter = "Файлы Excel (.xlsx) |*.xlsx| Файлы Excel 97-2003 (.xls) |*.xls";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.Multiselect = false;

            // Call the ShowDialog method to show the dialog box.
            DialogResult userClickedOK = openFileDialog1.ShowDialog();
			
            // Process input if the user clicked OK.
            if (userClickedOK == DialogResult.OK)
            {
            	this.excelFile = openFileDialog1.FileName;
            	this.textBox1.Text = openFileDialog1.SafeFileName;
            }	
		}
		void Button3Click(object sender, EventArgs e)
		{
			if((this.wordFile == "") || (this.excelFile == ""))
			{
				MessageBox.Show("Выберите Excel-данные и Word-шаблон!","Ошибка!", MessageBoxButtons.OK);
				return;
			}
			if (this.processExcel() && this.checkDictionary())
				this.processWord();
		}
		
		private bool processExcel() {
			// Results
			List<Student> students = new List<Student>();
			// Errors
			List<string> errors = new List<string>();
			// Progress
			ProgressorForm progress = null;
			// Excel
			Excel._Application excelApp = new Excel.Application();
			Excel.Workbook excelBook = excelApp.Workbooks.Add(this.excelFile);
			Excel._Worksheet workSheet = (Excel.Worksheet)(excelBook.ActiveSheet);
			int row = 1;
	    	int column = 1;
			try {
	    		row = 1;
	    		column = 1;
	    		while((workSheet.Cells[row,column] as Excel.Range).Value != null) row++;
	    		progress = new ProgressorForm(row);
	    		progress.Show();
	    		
	    		// Start
	    		row = 1;
	    		column = 1;
	    		while((workSheet.Cells[row,column] as Excel.Range).Value != null) {
	    			string fio = null;
	    			string group = null;
	    			try {
	    				fio = (workSheet.Cells[row,column] as Excel.Range).Value2.ToString();
	    			}catch(Exception e){
	    				errors.Add("Ошибка - не прочитана строка " + row.ToString() + " (первый столбец)!");
	    			}
	    			try {
	    				group = (workSheet.Cells[row,column +  1] as Excel.Range).Value2.ToString(); 
	    			}catch(Exception e){
	    				errors.Add("Ошибка - не прочитана строка " + row.ToString() + " (второй столбец)!");
	    			}
	    			
	    			if((fio != null) && (group != null)) {
	    				Student student = new Student();
	    				student.fio = fio;
	    				student.group = group;
	    				students.Add(student);
	    			}
	    			row++;
	    			progress.DoStep();
	    		}
	    		
			}catch(Exception e) {
	    		errors.Add("Ошибка чтения из файла! На строке: " + row.ToString() + e.ToString());
			}finally{
				if(excelBook != null)
					excelBook.Close();
	    		if(excelApp != null)
	    			excelApp.Quit();
	    		if(progress != null)
	    			progress.Quit();
			}
			if( errors.Count > 0 )
			{
				string msg = "";
				foreach(string key in errors) 
					msg += System.Environment.NewLine + key;
				MessageBox.Show(msg, "Ошибки при выгрузке", MessageBoxButtons.OK);
				return false;
			}
			this.fios = students;
			return true;
		}
		
		private bool checkDictionary() {
			List<string> errors = new List<string>();
			Regex reg = new Regex(@"^(\S+) (\S+)");
			foreach(Student fio in this.fios) {
				if( !reg.IsMatch(fio.fio) )
					errors.Add("Несоответсвие ФИО (" + fio.fio + ") шаблону: ФАМИЛИЯ ПРОБЕЛ ИМЯ ЛЮБОЙ-ТЕКСТ");
			}
			if(errors.Count > 0) {
				string msg = "";
				foreach(string err in errors)
					msg += System.Environment.NewLine + err;
				MessageBox.Show(msg, "Ошибки проверки синтаксиса", MessageBoxButtons.OK);
				return false;
			}
			return true;
		}
		
		private bool processWord() {
			Word.Application wordApp = null;
			Word.Document wordDoc = null;
			try {
				Object filename = this.wordFile;
				wordApp = new Microsoft.Office.Interop.Word.Application { Visible = true };
				wordDoc = wordApp.Documents.Open(ref filename, ReadOnly: false, Visible: true);
        		wordDoc.Activate();
			}catch(Exception e) {
				MessageBox.Show("Ошибка открытия файла Word! " + e.ToString(), "Ошибка открытия файла!", MessageBoxButtons.OK);
				if(wordDoc != null) 
					wordDoc.Close();
				if(wordApp != null)
					wordApp.Quit();
				return false;
			}
			
			// Замена
			try {
				
				foreach( Student student in this.fios ) {
						string[] fiostring = Regex.Split(student.fio,@" ");
						string family = fiostring[0];
						string name = fiostring[1];
						// Ф
		        		Word.Find fnd = wordApp.ActiveWindow.Selection.Find;
		        		fnd.ClearFormatting();
			        	fnd.Replacement.ClearFormatting();
			        	fnd.Forward = true;
			        	fnd.Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue;
			        	fnd.Text = "Фамилия";
			        	fnd.Replacement.Text = family;
			        	fnd.Execute(Replace: WdReplace.wdReplaceOne);
		        		// И
		        		fnd = wordApp.ActiveWindow.Selection.Find;
		        		fnd.ClearFormatting();
			        	fnd.Replacement.ClearFormatting();
			        	fnd.Forward = true;
			        	fnd.Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue;
			        	fnd.Text = "Имя";
			        	fnd.Replacement.Text = name;
			        	fnd.Execute(Replace: WdReplace.wdReplaceOne);
		        		// Группа
		        		fnd = wordApp.ActiveWindow.Selection.Find;
		        		fnd.ClearFormatting();
			        	fnd.Replacement.ClearFormatting();
			        	fnd.Forward = true;
			        	fnd.Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue;
			        	fnd.Text = "Нгруппа";
			        	fnd.Replacement.Text = student.group;
			        	fnd.Execute(Replace: WdReplace.wdReplaceOne);
				}
				// Close and save				
				wordDoc.Save();
		        wordDoc.Close();
		        wordDoc = null;
		        wordApp.Quit();
		        wordApp = null;
		        MessageBox.Show("Обработка бейджей выполнена успешно!","Выгрузка произошла успешно!", MessageBoxButtons.OK);
			}catch(Exception e){
				MessageBox.Show("Ошибка открытия файла Word! " + e.ToString(), "Ошибка открытия файла!", MessageBoxButtons.OK);
				return false;
			}finally{
				if(wordDoc != null) 
					wordDoc.Close();
				if(wordApp != null)
					wordApp.Quit();
			}
        	return true;
		}
	}
}
