using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace qldb
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btn_Export_Click(object sender, EventArgs e)
        {
            ExportReport();
        }

        public void ExportReport()
        {
            var excel = new Excel.Application();
            try
            {
                var obook = excel.Workbooks.Add();
                string excelName = "Bao cao thang 7.xlsx";
                var osheet = (Excel.Worksheet)obook.Sheets.Add();
                osheet.Name = "Thang 8";
                Excel.Range exRange;

                //osheet.SetColumnWidth(1, 5);
                Excel.Style style = obook.Styles.Add("style");
                style.Font.Name = "Times New Roman";

                #region Title and Datetime
                osheet.get_Range("A1", "R8").Font.Name = "Times New Roman"; //format font for report header 
                osheet.get_Range("A1", "R8").Font.Size = 14; //set font size 
                osheet.get_Range("H1", "R2").Font.Bold = true; // format chu in dam cho quoc hieu tieu ngu
                osheet.get_Range("A2", "G2").Font.Bold = true; // format in dam cho ten chi cuc
                osheet.get_Range("H3", "R3").Font.Italic = true; // format chu in nghieng cho ngay thang
                osheet.get_Range("A5", "R6").Font.Bold = true; // format chu in dam cho tieu de bao cao
                osheet.get_Range("A7", "R7").Font.Italic = true; // set chu in nghieng cho khoang thoi gian

                osheet.get_Range("A1", "G1").Merge(false);
                osheet.get_Range("A1", "G1").Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                osheet.Cells[1] = "CỤC QUẢN LÍ ĐƯỜNG BỘ II";

                osheet.get_Range("A2", "G2").Merge(false);
                exRange = osheet.get_Range("A2", "G2");
                exRange.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                exRange.FormulaR1C1 = "CHI CỤC QUẢN LÍ ĐƯỜNG BỘ II.6";

                osheet.get_Range("A3", "G3").Merge(false);
                exRange = osheet.get_Range("A3", "G3");
                exRange.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                exRange.FormulaR1C1 = "Số:  163 /BC-CCQLĐBII.6";   

                //quoc hieu tieu ngu
                osheet.get_Range("H1", "R1").Merge(false);
                exRange = osheet.get_Range("H1", "R1");
                exRange.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                exRange.FormulaR1C1 = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM";
                
                osheet.get_Range("H2", "R2").Merge(false);
                exRange = osheet.get_Range("H2", "R2");
                exRange.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                exRange.FormulaR1C1 = "Độc lập - Tự do - Hạnh phúc";
                // end quoc hieu tieu ngu

                osheet.get_Range("H3", "R3").Merge(false);
                exRange = osheet.get_Range("H3", "R3");
                exRange.Cells.HorizontalAlignment = XlHAlign.xlHAlignRight;
                exRange.FormulaR1C1 = "Thừa Thiên Huế, ngày 15 tháng 2 năm 2019";

                osheet.get_Range("A5", "R5").Merge(false);
                exRange = osheet.get_Range("A5", "R5");
                exRange.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                exRange.FormulaR1C1 = "BÁO CÁO";

                osheet.get_Range("A6", "R6").Merge(false);
                exRange = osheet.get_Range("A6", "R6");
                exRange.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                exRange.FormulaR1C1 = "TỔNG HỢP TAI NẠN GIAO THÔNG ĐƯỜNG BỘ THÁNG 8 NĂM 2018";

                osheet.get_Range("A7", "R7").Merge(false);
                exRange = osheet.get_Range("A7", "R7");
                exRange.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                exRange.FormulaR1C1 = "(Từ ngày 15/7/2019 đến ngày 15/8/2019)";

                osheet.get_Range("A8", "R8").Merge(false);
                exRange = osheet.get_Range("A8", "R8");
                exRange.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                exRange.FormulaR1C1 = "Kính gửi: Cục Quản lí đường bộ II";
                #endregion

                #region table Header

                osheet.get_Range("A9", "R11").Cells.Font.Bold = true; //Format bold for text of table header
                osheet.get_Range("A9", "R11").Cells.Font.Size = 12; //Set font size for text of table header
                osheet.get_Range("A9", "R11").Cells.Borders.LineStyle = XlLineStyle.xlContinuous; // draw border for table header
                osheet.get_Range("A9", "R11").Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                osheet.get_Range("A9", "R11").Cells.VerticalAlignment = XlVAlign.xlVAlignCenter;
                osheet.get_Range("A9", "R11").WrapText = true;
                osheet.get_Range("A9", "R11").Font.Name = "Times New Roman";

                //STT
                osheet.get_Range("A9", "A11").Merge(false); // merge cells
                exRange = osheet.get_Range("A9", "A11");
                exRange.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter; // format center text horizontal
                exRange.Cells.VerticalAlignment = XlVAlign.xlVAlignCenter; // format center for text vertical
                exRange.FormulaR1C1 = "TT"; // Set text for cell
                exRange.ColumnWidth = 4; // set width for cell
                //end STT

                //TNGT duong bo
                osheet.get_Range("B9", "D11").Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                osheet.get_Range("B9", "D11").Cells.VerticalAlignment = XlVAlign.xlVAlignCenter;
                osheet.get_Range("B9", "D11").WrapText = true;

                osheet.get_Range("B9", "D10").Merge(false);
                exRange = osheet.get_Range("B9", "D10");
                exRange.FormulaR1C1 = "TNGT ĐƯỜNG BỘ";
                

                osheet.get_Range("B11").FormulaR1C1 = "Tên đường";

                osheet.get_Range("C11").FormulaR1C1 = "Vị trí, lí trình, thời gian";
                osheet.get_Range("C11").ColumnWidth = 23;

                osheet.get_Range("D11").FormulaR1C1 = "Tổng số vụ xảy ra trong tháng";
                osheet.get_Range("D11").ColumnWidth = 14;
                // end TNGT duong bo

                //Nguyen nhan xay ra
                osheet.get_Range("E9", "G10").Merge(false);
                exRange = osheet.get_Range("E9", "G10");
                exRange.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                exRange.FormulaR1C1 = "Nguyên nhân xảy ra";

                osheet.get_Range("E11").FormulaR1C1 = "Do đường";
                osheet.get_Range("F11").FormulaR1C1 = "Do người";
                osheet.get_Range("G11").FormulaR1C1 = "Do phương tiện";
                osheet.get_Range("G11").ColumnWidth = 45;
                //end Nguyen nhan xay ra

                //Thiet hai
                osheet.get_Range("H9", "K9").Merge(false);
                exRange = osheet.get_Range("H9", "K9");
                exRange.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                exRange.FormulaR1C1 = "Thiệt hại";

                osheet.get_Range("H10", "I10").Merge(false);
                exRange = osheet.get_Range("H10", "I10");
                exRange.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                exRange.FormulaR1C1 = "Số người";

                osheet.get_Range("J10", "K10").Merge(false);
                exRange = osheet.get_Range("J10", "K10");
                exRange.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                exRange.FormulaR1C1 = "Giá trị(triệu đồng)";

                osheet.get_Range("H11").FormulaR1C1 = "Chết";
                osheet.get_Range("I11").FormulaR1C1 = "Bị thương";
                osheet.get_Range("I11").ColumnWidth = 12;

                osheet.get_Range("J11").FormulaR1C1 = "Cầu, đường";
                osheet.get_Range("J11").ColumnWidth = 16;

                osheet.get_Range("K11").FormulaR1C1 = "Phương tiện";
                osheet.get_Range("K11").ColumnWidth = 11;
                //end thiet hai

                //so sanh
                osheet.get_Range("L9", "N10").Merge(false);
                exRange = osheet.get_Range("L9", "N10");
                exRange.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                exRange.FormulaR1C1 = "So sánh với tháng liền kề";
                exRange.Style.WrapText = true;

                osheet.get_Range("L11").FormulaR1C1 = "Tổng số vụ xảy ra";
                osheet.get_Range("L11").ColumnWidth = 12;

                osheet.get_Range("M11").FormulaR1C1 = "Số người chết";
                osheet.get_Range("M11").ColumnWidth = 11;

                osheet.get_Range("N11").FormulaR1C1 = "Số người bị thương";
                osheet.get_Range("N11").ColumnWidth = 12;
                //end so sanh

                //tong hop
                osheet.get_Range("O9", "Q10").Merge(false);
                exRange = osheet.get_Range("O9", "Q10");
                exRange.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                exRange.FormulaR1C1 = "Tổng hợp từ đầu năm đến tháng báo cáo";
                //exRange.Columns.AutoFit();

                osheet.get_Range("O11").FormulaR1C1 = "Tổng số vụ xảy ra";
                osheet.get_Range("P11").FormulaR1C1 = "Số người chết";
                osheet.get_Range("P11").ColumnWidth = 10;

                osheet.get_Range("Q11").FormulaR1C1 = "Số người bị thương";
                osheet.get_Range("Q11").ColumnWidth = 11;
                //end tong hop
                //ghi chu
                osheet.get_Range("R9", "R11").Merge(false);
                exRange = osheet.get_Range("R9", "R10");
                exRange.Cells.VerticalAlignment = XlVAlign.xlVAlignCenter;
                exRange.FormulaR1C1 = "Ghi chú";
                osheet.get_Range("R11").ColumnWidth = 12;
                #endregion

                #region table body category

                osheet.get_Range("A12", "R12").Font.Name = "Times New Roman";
                osheet.get_Range("A12", "R12").Font.Size = 12;
                osheet.get_Range("A12", "K12").Font.Bold = true;
                osheet.get_Range("O12", "Q12").Font.Bold = true;
                osheet.get_Range("A12", "R12").WrapText = true;
                osheet.get_Range("A12", "R12").HorizontalAlignment = XlHAlign.xlHAlignCenter;
                osheet.get_Range("A12", "R12").VerticalAlignment = XlVAlign.xlVAlignCenter;
                osheet.get_Range("A12", "R12").Borders.LineStyle = XlLineStyle.xlContinuous;

                osheet.get_Range("A12").FormulaR1C1 = "I";//TT
                osheet.get_Range("B12", "C12").Merge(false);
                osheet.get_Range("B12", "C12").FormulaR1C1 = "Quốc lộ 1";
                osheet.get_Range("B12", "C12").HorizontalAlignment = XlHAlign.xlHAlignLeft;

                osheet.get_Range("D12").FormulaR1C1 = "QL.1"; //tong so vu xay ra trong thang
                osheet.get_Range("E12").FormulaR1C1 = "QL.1"; // do duong   
                osheet.get_Range("F12").FormulaR1C1 = "QL.1"; //do nguoi    
                osheet.get_Range("G12").FormulaR1C1 = "12"; // do phuong tien
                osheet.get_Range("H12").FormulaR1C1 = "QL.1"; //chet
                osheet.get_Range("I12").FormulaR1C1 = "QL.1"; //bi thuong
                osheet.get_Range("J12").FormulaR1C1 = "QL.1"; //cau, duong
                osheet.get_Range("K12").FormulaR1C1 = "QL.1"; //phuong tien
                osheet.get_Range("L12").FormulaR1C1 = "Tăng 05 vụ"; // tong so vu xay ra
                osheet.get_Range("M12").FormulaR1C1 = "QL.1"; // so nguoi chet
                osheet.get_Range("N12").FormulaR1C1 = "Giảm 05 người"; // so nguoi bị thuong 
                osheet.get_Range("O12").FormulaR1C1 = "QL.1"; // //tong so vu dau năm den thang bao cao
                osheet.get_Range("P12").FormulaR1C1 = "QL.1"; // so nguoi chet (dau năm)
                osheet.get_Range("Q12").FormulaR1C1 = "QL.1"; // so nguoi bị thuong (dau nam - thang bao cao)
                osheet.get_Range("R12").FormulaR1C1 = "QL.1"; // GHI CHu
                #endregion

                #region table body details
                string a = "13";
                osheet.get_Range("A13", "R13").Borders.LineStyle = XlLineStyle.xlContinuous;
                osheet.get_Range("A13", "R13").Font.Name = "Times New Roman";
                osheet.get_Range("A13", "R13").Font.Size = 13;
                osheet.get_Range("A13", "R13").WrapText = true;
                osheet.get_Range("A13", "R13").HorizontalAlignment = XlHAlign.xlHAlignCenter;
                osheet.get_Range("A13", "R13").VerticalAlignment = XlVAlign.xlVAlignCenter;
                osheet.get_Range("G13").HorizontalAlignment = XlHAlign.xlHAlignGeneral;
                osheet.Range["A13:R13"].Interior.Color = Excel.XlRgbColor.rgbYellow;

                osheet.get_Range("A" + a).FormulaR1C1 = "1"; //TT
                osheet.get_Range("B13").FormulaR1C1 = "QL.1"; //Ten duong
                osheet.get_Range("C13").FormulaR1C1 = "QL.1"; // vi tri
                osheet.get_Range("D13").FormulaR1C1 = "QL.1"; //tong so vu xay ra trong thang
                osheet.get_Range("E13").FormulaR1C1 = "QL.1"; // do duong   
                osheet.get_Range("F13").FormulaR1C1 = "QL.1"; //do nguoi    
                osheet.get_Range("G13").FormulaR1C1 = "Xe ô tô tải 79C - 087.66 chạy hướng Nam - Bắc không làm chủ tốc độ đã va chạm với người đi bộ đang qua đường gây tai nạn"; // do phuong tien
                osheet.get_Range("H13").FormulaR1C1 = "QL.1"; //chet
                osheet.get_Range("I13").FormulaR1C1 = "QL.1"; //bi thuong
                osheet.get_Range("J13").FormulaR1C1 = "QL.1"; //cau, duong
                osheet.get_Range("K13").FormulaR1C1 = "QL.1"; //phuong tien
                osheet.get_Range("L13").FormulaR1C1 = "QL.1"; // tong so vu xay ra
                osheet.get_Range("M13").FormulaR1C1 = "QL.1"; // so nguoi chet
                osheet.get_Range("N13").FormulaR1C1 = "QL.1"; // so nguoi bị thuong 
                osheet.get_Range("O13").FormulaR1C1 = "QL.1"; // //tong so vu dau năm den thang bao cao
                osheet.get_Range("P13").FormulaR1C1 = "QL.1"; // so nguoi chet (dau năm)
                osheet.get_Range("Q13").FormulaR1C1 = "QL.1"; // so nguoi bị thuong (dau nam - thang bao cao)
                osheet.get_Range("R13").FormulaR1C1 = "QL.1"; // GHI CHu

                #endregion

                #region footer
                int vitri = 15;
                //string end = (int.Parse(vitri) + 11).ToString();
                //footer left 
                osheet.get_Range("A" + vitri, "R" + (vitri + 11)).Font.Name = "Times New Roman";
                osheet.get_Range("A" + vitri, "R" + (vitri + 11)).Font.Size = 13;

                osheet.get_Range("A" + vitri, "F" + vitri).Merge(false);
                osheet.get_Range("A" + vitri, "F" + vitri).FormulaR1C1 = "Nhận xét và kiến nghị các vị trí hay xảy ra tai nạn:";

                osheet.get_Range("A" + (vitri + 1), "C" + (vitri + 1)).Merge(false);
                osheet.get_Range("A" + (vitri + 1), "C" + (vitri + 1)).Font.Bold = true;
                osheet.get_Range("A" + (vitri + 1), "C" + (vitri + 1)).Font.Italic = true;
                osheet.get_Range("A" + (vitri + 1), "C" + (vitri + 1)).FormulaR1C1 = " Nơi nhận:";

                osheet.get_Range("A" + (vitri + 2), "B" + (vitri + 2)).Merge(false);
                osheet.get_Range("A" + (vitri + 2), "B" + (vitri + 2)).FormulaR1C1 = " - Như trên;";

                osheet.get_Range("A" + (vitri + 3), "C" + (vitri + 3)).Merge(false);
                osheet.get_Range("A" + (vitri + 3), "C" + (vitri + 3)).FormulaR1C1 = " - Phòng ATGT,QLBT(b/c);";

                osheet.get_Range("A" + (vitri + 4), "B" + (vitri + 4)).Merge(false);
                osheet.get_Range("A" + (vitri + 4), "B" + (vitri + 4)).FormulaR1C1 = " - Lưu: VT.";
                //end footer left 

                //footer right
                osheet.get_Range("L" + (vitri + 1), "R" + (vitri + 11)).Font.Bold = true;
                osheet.get_Range("L" + (vitri + 1), "R" + (vitri + 11)).HorizontalAlignment = XlHAlign.xlHAlignCenter;

                osheet.get_Range("L" + (vitri + 1), "R" + (vitri + 1)).Merge(false);
                osheet.get_Range("L" + (vitri + 1), "R" + (vitri + 1)).FormulaR1C1 = "KT. CHI CỤC TRƯỞNG";

                osheet.get_Range("L" + (vitri + 2), "R" + (vitri + 2)).Merge(false);
                osheet.get_Range("L" + (vitri + 2), "R" + (vitri + 2)).FormulaR1C1 = "PHÓ CHI CỤC TRƯỞNG";

                osheet.get_Range("L" + (vitri + 11), "R" + (vitri + 11)).Merge(false);
                osheet.get_Range("L" + (vitri + 11), "R" + (vitri + 11)).FormulaR1C1 = "Đặng Nguyễn Ngọc Linh";
                //end footer right
                #endregion
                //save and export file excel 
                obook.SaveAs(@"C:\Users\DevyDuong\Desktop\" + excelName);
                obook.Close();
                excel.Quit();
                GC.Collect();
            }
            catch(Exception)
            {
                excel.Quit();
            }
        }
    }
}
