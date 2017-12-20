using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using NPOI.XWPF.UserModel;
using SoftFluent.Windows;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace NOPI_Demo
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

        }

        private void exportXLSX_Click(object sender, RoutedEventArgs e)
        {
            var newFile = @"newbook.core.xlsx";

            using (var fs = new FileStream(newFile, FileMode.Create, FileAccess.Write))
            {
                // XSSFWorkbook : *.xlsx >= Excel2007
                // HSSFWorkbook : *.xls  < Excel2007
                IWorkbook workbook = new XSSFWorkbook();

                ISheet sheet1 = workbook.CreateSheet("Sheet Name");

                // 所有索引都从0开始

                // 合并单元格
                sheet1.AddMergedRegion(new CellRangeAddress(0, 0, 0, 10));

                var rowIndex = 0;
                IRow row = sheet1.CreateRow(rowIndex); //创建行
                row.Height = 30 * 80;
                row.CreateCell(0).SetCellValue("this is content");
                sheet1.AutoSizeColumn(0); //按照值的长短 自动调节列的大小
                rowIndex++;

                // 插入图片
                byte[] data = File.ReadAllBytes(@"image.jpg");
                int picInd = workbook.AddPicture(data, NPOI.SS.UserModel.PictureType.JPEG);
                XSSFCreationHelper helper = workbook.GetCreationHelper() as XSSFCreationHelper;
                XSSFDrawing drawing = sheet1.CreateDrawingPatriarch() as XSSFDrawing;
                XSSFClientAnchor anchor = helper.CreateClientAnchor() as XSSFClientAnchor;
                anchor.Col1 = 10;
                anchor.Row1 = 0;
                XSSFPicture pict = drawing.CreatePicture(anchor, picInd) as XSSFPicture;
                pict.Resize();

                // 新建sheet
                var sheet2 = workbook.CreateSheet("Sheet2");
                // 更改样式
                var style1 = workbook.CreateCellStyle();
                style1.FillForegroundColor = HSSFColor.Blue.Index2;
                style1.FillPattern = FillPattern.SolidForeground;

                var style2 = workbook.CreateCellStyle();
                style2.FillForegroundColor = HSSFColor.Yellow.Index2;
                style2.FillPattern = FillPattern.SolidForeground;

                var cell2 = sheet2.CreateRow(0).CreateCell(0);
                cell2.CellStyle = style1;
                cell2.SetCellValue(0);

                cell2 = sheet2.CreateRow(1).CreateCell(0);
                cell2.CellStyle = style2;
                cell2.SetCellValue(1);


                //保存
                workbook.Write(fs);
            }
            txtStatus.Text = "writing xlsx successful!";
        }

        private void exportDOCX_Click(object sender, RoutedEventArgs e)
        {
            var newFile2 = @"newbook.core.docx";
            using (var fs = new FileStream(newFile2, FileMode.Create, FileAccess.Write))
            {
                XWPFDocument doc = new XWPFDocument();
                var p0 = doc.CreateParagraph();
                p0.Alignment = ParagraphAlignment.CENTER;
                XWPFRun r0 = p0.CreateRun();
                r0.FontFamily = "microsoft yahei";
                r0.FontSize = 18;
                r0.IsBold = true;
                r0.SetText("This is title");

                var p1 = doc.CreateParagraph();
                p1.Alignment = ParagraphAlignment.LEFT;
                p1.IndentationFirstLine = 500;
                XWPFRun r1 = p1.CreateRun();
                r1.FontFamily = "宋体";
                r1.FontSize = 12;
                r1.IsBold = true;
                r1.SetText("中文宋体加粗This is content, content content content content content content content content content");

                doc.Write(fs);
            }
            txtStatus.Text = "writing docx successful!";
        }
    }
}
