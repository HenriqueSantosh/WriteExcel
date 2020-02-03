using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;

namespace TESTELINK
{
    class Program
    {
        static void Main(string[] args)
        {
            IList<Conta> ListConta = new List<Conta> { new Conta() { IConta = 200, IdBovespa = 2356 },
            new Conta() { IConta = 220, IdBovespa = 2356 },
            new Conta() { IConta = 210, IdBovespa = 2336 },
            new Conta() { IConta = 230, IdBovespa = 23596 },
             new Conta() { IConta = 500000, IdBovespa = 23596 }};


            IList<ContaDetalha> contaDetalhas = new List<ContaDetalha> {
             new ContaDetalha { ContaNumero = 200,DEP ="DIP",Valor = 3000},
             new ContaDetalha {ContaNumero = 220, DEP = "DEP",Valor = 2500 },
             new ContaDetalha { ContaNumero = 210,DEP = "DIP", Valor = 1500},
             new ContaDetalha { ContaNumero = 230,DEP="DIP",Valor = 3500},
             new ContaDetalha { ContaNumero = 500,DEP="DIP",Valor = 3500}};
            //left join com linq
            var listNewClasse = (from contaN in ListConta
                                 join contaDet in contaDetalhas
                                  on contaN.IConta equals contaDet.ContaNumero
                                  into A
                                 from B in A.DefaultIfEmpty(new ContaDetalha())
                                 select
                                 new NovaClasse()
                                 {
                                     ContaNumero = contaN.IConta,
                                     IdBovespa = contaN.IdBovespa,
                                     DEP = B.DEP == (null) ? "N/A" : B.DEP,
                                     Valor = B.DEP == null ? 0 : B.DEP.Equals("DIP") ? B.Valor : -B.Valor
                                 }).ToList();
            writeExcel(listNewClasse);
            Console.ReadKey();

        }

        public static void writeExcel(IList<NovaClasse> novaClasses)
        {
            try
            {
                using (var excelPackage = new ExcelPackage())
                {
                    excelPackage.Workbook.Properties.Author = "UBS Brasil";
                    excelPackage.Workbook.Properties.Title = "UBs BRAsil";

                    var sheet = excelPackage.Workbook.Worksheets.Add("Planilha 1");
                    sheet.Name = "Ubs teste";

                    // Títulos
                    var i = 1;
                    PropertyInfo[] propriedades = typeof(NovaClasse).GetProperties();
                    var titulos = new String[] { "Título Um", "Título Dois", "Título Três" };
                    foreach (var titulo in propriedades)
                    {
                        sheet.Cells[1, i].Style.Font.Bold = true;
                        sheet.Cells[1, i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[1, i].Style.Fill.BackgroundColor.SetColor(Color.Red);
                        sheet.Cells[1, i++].Value = titulo.Name;
                    }

                    // Valores
                    i = 1;
                    int column = 2;
                    foreach (var valor in novaClasses)
                    {

                        // Aqui escrevo a segunda linha do arquivo com alguns valores.
                        sheet.Cells[column, i++].Value = valor.ID;
                        sheet.Cells[column, i++].Value = valor.IdBovespa;
                        sheet.Cells[column, i++].Value = valor.ContaNumero;
                        sheet.Cells[column, i++].Value = valor.DEP;
                        sheet.Cells[column, i++].Value = valor.Valor;
                        i = 1;
                        column++;

                    }

                    sheet.Cells[column++, i].Value = "Total";
                    sheet.Cells[column, i].Value = novaClasses.Sum(total => total.Valor);

                    string path = @"C:\Users\Henrique\teste.xlsx";
                    File.WriteAllBytes(path, excelPackage.GetAsByteArray());
                }
                Console.WriteLine("executado com sucesso");

            }
            catch (Exception ex)
            {
                Console.WriteLine("Ocorreu u erro ao executar " + ex.Message.ToString());
            }
            
            }
        }
    }

