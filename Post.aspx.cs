using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using System.Configuration;
using System.Data.Odbc;

public partial class _Post : System.Web.UI.Page
{
	 string onlinePDF="C:\\IIS\\ImageDB_RPT\\onlinepdf\\";

    protected void Page_Load(object sender, EventArgs e)
    {
        string ODBCName = "post";
        //string DBIP = "localhost";
        string DBName = "post"; 
        string UserID = "root";
        string UserPassword = "12345678";   
        
        
        try
        {                     
            string rptFile = "";
            ReportDocument report = new ReportDocument();
            CrystalReportViewer1.DisplayGroupTree  = false;
            
            int i = Convert.ToInt16(Request["rpt"]);
            string randomId=Request["randomId"];
            switch (i)
            {
                case 1:
                    rptFile = this.Server.MapPath("post/report1.rpt");
                    report.Load(rptFile);                    
                    report.SetDatabaseLogon(UserID, UserPassword, ODBCName, DBName);
                    //report.SetParameterValue("customerId", Request["customerId"] );
                    //saveDisk1(report, "CustomerPriceList"); 
                break; 
                
                 case 2:
                    rptFile = this.Server.MapPath("post/report2.rpt");
                    report.Load(rptFile);                    
                    report.SetDatabaseLogon(UserID, UserPassword, ODBCName, DBName);
                    //report.SetParameterValue("customerId", Request["customerId"] );
                    //saveDisk1(report, "CustomerPriceList"); 
                break; 
                
                 case 3:
                    rptFile = this.Server.MapPath("post/report3.rpt");
                    report.Load(rptFile);                    
                    report.SetDatabaseLogon(UserID, UserPassword, ODBCName, DBName);
                    //report.SetParameterValue("customerId", Request["customerId"] );
                    //saveDisk1(report, "CustomerPriceList"); 
                break; 
                
                 case 4:
                    rptFile = this.Server.MapPath("post/report4.rpt");
                    report.Load(rptFile);                    
                    report.SetDatabaseLogon(UserID, UserPassword, ODBCName, DBName);
                    //report.SetParameterValue("customerId", Request["customerId"] );
                    //saveDisk1(report, "CustomerPriceList"); 
                break; 
                
                 case 5:
                    rptFile = this.Server.MapPath("post/report5.rpt");
                    report.Load(rptFile);                    
                    report.SetDatabaseLogon(UserID, UserPassword, ODBCName, DBName);
                    //report.SetParameterValue("customerId", Request["customerId"] );
                    //saveDisk1(report, "CustomerPriceList"); 
                break; 
                
            }
           
            CrystalReportViewer1.ReportSource = report;
            CrystalReportViewer1.DataBind();            
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine("printReport=" + ex.Message);
        }
        
    }
    

  public void SaveStreamToFile(string fileFullPath, System.IO.Stream stream)
   {
    if (stream.Length == 0) return;
    // Create a FileStream object to write a stream to a file
    using (System.IO.FileStream fileStream = System.IO.File.Create(fileFullPath, (int)stream.Length))
    {
        // Fill the bytes[] array with the stream data
        byte[] bytesInStream = new byte[stream.Length];
        stream.Read(bytesInStream, 0, (int)bytesInStream.Length);
        // Use FileStream object to write to the specified file
        fileStream.Write(bytesInStream, 0, bytesInStream.Length);
     }
   }

   public void saveDisk1(ReportDocument report, String fileName)
    {       
                    
                System.IO.Stream stream1 = report.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat);
                SaveStreamToFile(onlinePDF+"\\"+fileName+".pdf", stream1);
    }
    
  
        
   public void saveDisk(ReportDocument report, String fileName, String fileType)
    {       
        switch (fileType)
        {
            case "excel":
                System.IO.Stream stream = report.ExportToStream(CrystalDecisions.Shared.ExportFormatType.Excel);
                byte[] bytes = new byte[stream.Length];
                stream.Read(bytes, 0, bytes.Length);
                stream.Seek(0, System.IO.SeekOrigin.Begin);
                //export file
                Response.ClearContent();
                Response.ClearHeaders();
                Response.AddHeader("content-disposition", "attachment;filename=" + fileName+".xls");
                Response.ContentType = "application/vnd.ms-excel";
                Response.OutputStream.Write(bytes, 0, bytes.Length);
                Response.Flush();
                Response.Close();
            break;
            case "pdf":
                System.IO.Stream stream1 = report.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat);
                byte[] bytes1 = new byte[stream1.Length];
                stream1.Read(bytes1, 0, bytes1.Length);
                stream1.Seek(0, System.IO.SeekOrigin.Begin);
                //export file
                Response.ClearContent();
                Response.ClearHeaders();
                Response.AddHeader("content-disposition", "attachment;filename=" + fileName+".pdf");
                Response.ContentType = "application/pdf";
                Response.OutputStream.Write(bytes1, 0, bytes1.Length);
                Response.Flush();
                Response.Close();
                break;
            case "word":
                System.IO.Stream stream2 = report.ExportToStream(CrystalDecisions.Shared.ExportFormatType.WordForWindows);
                byte[] bytes2 = new byte[stream2.Length];
                stream2.Read(bytes2, 0, bytes2.Length);
                stream2.Seek(0, System.IO.SeekOrigin.Begin);
                //export file
                Response.ClearContent();
                Response.ClearHeaders();
                Response.AddHeader("content-disposition", "attachment;filename=" + fileName+".doc");
                Response.ContentType = "application/vnd.ms-word";
                Response.OutputStream.Write(bytes2, 0, bytes2.Length);
                Response.Flush();
                Response.Close();
            break;
        }        
    }
    
}