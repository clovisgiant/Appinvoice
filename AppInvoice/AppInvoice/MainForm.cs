/*
 * Criado por SharpDevelop.
 * Usuário: clovis
 * Data: 17/07/2013
 * Hora: 17:52
 * 
 * Para alterar este modelo use Ferramentas | Opções | Codificação | Editar Cabeçalhos Padrão.
 */
using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Data;
using System.Configuration;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Collections;
using System.IO;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Text;
using System.Globalization;




namespace AppInvoice
{
	/// <summary>
	/// Description of MainForm.
	/// </summary>
	public partial class MainForm : Form
	{
		
		
		public string Ano;
        public string Processo;	
        public string menssagem;
        public string Invoice;
		public string Outbox;        
       
		
		public MainForm()
		{
			//
			// The InitializeComponent() call is required for Windows Forms designer support.
			//
			InitializeComponent();
			 	
			
			//
			// TODO: Add constructor code after the InitializeComponent() call.
			//
		}
		
		void MainFormLoad(object sender, EventArgs e)
		{
			
			string[] args = Environment.GetCommandLineArgs();
			//MessageBox.Show(args.Length.ToString());
			if (args.Length == 5)
			{
				
				Ano = args[1];
				Processo = args[2];
				Invoice =args[3];
				Outbox =args[4];				
				GeraInvoice( Ano,Processo,Invoice);
			}
			
			
	 				 
		}	
		
		
		public void GeraInvoice(string Ano, string Processo,string invoice)
		{
			SqlConnection conn;
			conn = dbConecta();		
			var rs = new IBM();
			
			  string query =" m_IBMExpRelInvoice '" + Ano + "', '" + Processo + "', '" + Invoice + "'";
			
			  SqlCommand cmd = new SqlCommand(query, conn);
			
			  try
			
			  {
			
			    conn.Open();
			
			    using (SqlDataReader rdr = cmd.ExecuteReader())
			
			    {
			
			      while (rdr.Read())
			
			      { 
			      	//var rs = new IBM();
			
				    rs.DataFatura = rdr[0].ToString();			
				    rs.TpInvoice = rdr[1].ToString();				    			
				    rs.invTO = rdr[2].ToString();
				    rs.ShipTO = rdr[3].ToString();			
				    rs.CaixaCAB = rdr[4].ToString();
				    rs.InvCTY = rdr[5].ToString();
				    rs.InvLoc =rdr[6].ToString();
				    rs.ShipCTY = rdr[7].ToString();			
				    rs.ShipLoc = rdr[8].ToString();
				    rs.InvToCty = rdr[9].ToString();
					rs.InvToLoc = rdr[10].ToString();				    
				    rs.ShipToCty = rdr[11].ToString();  
				    rs.ShipToLoc = rdr[12].ToString();				    
				    rs.Transp = rdr[13].ToString(); 
				    rs.DTerms =rdr[14].ToString();
				    rs.PedORder = rdr[15].ToString();   			
				    rs.ATTN = rdr[16].ToString(); 
				    rs.EMERG = rdr[17].ToString();
				    rs.Ident = rdr[18].ToString();				    
					rs.NroArquivo = rdr[34].ToString();
					rs.NomeArq = rdr[35].ToString();
					rs.Observacoes =rdr[36].ToString();		
					rs.Serial =rdr[37].ToString();						
					
			

			
			      }
			
			    }
			
			  }
			
			  catch(Exception ex)
			
			  {
			
			    MessageBox.Show(ex.Message);
			
			  }
			  
			conn.Close();           
			PdfPTable Itens;					
			var documento = new Document(PageSize.A4, 5, 5, 15, 15);
			var writer = PdfWriter.GetInstance(documento, new				                                   
			FileStream(Outbox+"\\"+rs.NomeArq+".pdf", FileMode.Create));					
			documento.Open();
			Itens = GetFormatItens(conn,documento,writer,rs);			
			documento.Close();				
			//MessageBox.Show("Invoice gerada em C:\\multproc\\outbox\\InvoicePDF.pdf");			
			System.Diagnostics.Process.Start(Outbox+"\\"+rs.NomeArq+".pdf");
		    //using the start method of system.diagnostics.process class
		    this.Close();
		}
		
		private void SetCabecalho(Document documento, Int32 TipoCab, PdfWriter writer,IBM rs)
		{
			BaseFont bfTimes = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, false);
		    BaseFont bfTimes2 = BaseFont.CreateFont(BaseFont.TIMES_BOLD, BaseFont.CP1252, false);
			BaseFont bfTimes3 = BaseFont.CreateFont(BaseFont.COURIER_BOLD, BaseFont.CP1252, false);
			BaseFont Helvetica = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, false);
			BaseFont HelveticaBold = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, false);
			BaseFont HelveticaBold2 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, false);
		    
			Font fontHelvetica = new Font(Helvetica, 10, 2 ,GrayColor.BLACK);
		    Font fontHelveticaBold = new Font(HelveticaBold, 10,0 ,GrayColor.BLACK);
		    Font fontHelveticaBold2 = new Font(HelveticaBold, 12,0 ,GrayColor.BLACK);

		    var times = new Font(bfTimes, 8, 2 ,GrayColor.BLUE);
			Font times2 = new Font(bfTimes,11, 2 ,GrayColor.BLACK);
			Font times3 = new Font(bfTimes, 8, 2 ,GrayColor.GRAYWHITE);
			
			Font arial = FontFactory.GetFont("Arial", 9, GrayColor.GRAYWHITE);
			Font arial2 = FontFactory.GetFont("Arial", 7, GrayColor.BLACK);			
			
			if (TipoCab == 1)
			{
				
				LogoDocumento(documento,writer);
				tabelaIvoiceTipoCab1(documento,writer,rs);
				SetColunns(writer,documento,rs);
								
			}
			else
			{	
				documento.NewPage();				
				LogoDocumento(documento,writer);
				tabelaIvoiceTipoCab2(documento,writer,rs);
				SetColunns2(writer,documento,rs);				
							
								           	 
			}
			//SetColunns(writer,documento);
			
			
		}
		
		private SqlConnection dbConecta()
		{	
			SqlConnection cno;
			string connect = "Data Source=SAOLA2K220;Integrated Security=SSPI;Initial Catalog=rjOperacional";
			//string connect = "Data Source=multsp3\\treinamento;Integrated Security=SSPI;Initial Catalog=dbOperacional";
			cno = new SqlConnection(connect);
			try {
				
			} catch (Exception ex) {
				
				MessageBox.Show(ex.Message);
			}
			return cno;

		}
		private PdfPTable GetFormatItens(SqlConnection conn, Document documento,PdfWriter writer, IBM rs)
		{
			
			
			Font Mixed = FontFactory.GetFont("Mixed", 9);
			Font font8 = FontFactory.GetFont("ARIAL", 8);														//'" + 2010 + "', '" + 27041980 + "', '" + 90428574 + "'"
			string query = " m_IBMExpRelInvoice '" + Ano + "', '" + Processo + "', '" + Invoice + "'";
			//string query = "select  ITPTN,ITDSC ,iTTEc,Unidade,ITQTD,Agrupamento,itpru,itnro From Itens where ano ='"+Ano+"' and processo ='"+Processo+"' and ftnro ='"+Invoice+ "'";
			//string query = "select  ITPTN,ITDSC ,iTTEc,Unidade,ITQTD,Agrupamento,itpru,itnro From Itens where ano ='2010'and processo ='27041980' and ftnro ='90428574' ";
			SqlDataReader Itens;			
			int contador = 0;
			int contadorItens =0;			
			int contadorLinhas =30;
			string TotalAMT ="";
			double Frete=0;
			string Ctype ="";
			string dterms="";
			string unit ="";
			string terms="";
			
			List<string> linhas;
			conn.Open();				
		    SqlCommand cmd = new SqlCommand(query, conn); 
					
			Itens = cmd.ExecuteReader();
			bool yesprimeira = true;
			var table = new PdfPTable(2) { WidthPercentage = 100}; var colWidthPercentages2 = new[] {9f,100f }; table.SetWidths(colWidthPercentages2);
			var table2 = new PdfPTable(9) { WidthPercentage = 100}; var colWidthPercentages3 = new[] { 10f, 30f, 10f, 10f,10f,10f,15f,15f,10f }; table2.SetWidths(colWidthPercentages3);
			var table3 = new PdfPTable(9) { WidthPercentage = 100}; var colWidthPercentages4 = new[] { 10f, 30f, 10f, 10f,10f,10f,15f,15f,10f }; table3.SetWidths(colWidthPercentages4);
				table.DefaultCell.Border = 0;
				table2.DefaultCell.Border = 0;
				table3.DefaultCell.Border = 0;					
				//contador++;
				while (Itens.Read())
					 	
				
				  { 
					  	//contador++;
											  	
							
						PdfPCell testetableCASE = new PdfPCell(new Phrase("CASE#",Mixed));//Colunas									
						testetableCASE.Border = 0;
						testetableCASE.HorizontalAlignment = 0;
						//testetableCASE.PaddingTop = 4f;
						
						PdfPCell testetablevazia = new PdfPCell(new Phrase("    ",Mixed));//Colunas									
						testetablevazia.Border = 0;
						testetablevazia.HorizontalAlignment =0;
						
						PdfPCell testetable = new PdfPCell(new Phrase(Itens[20].ToString(),Mixed));//Colunas									
						testetable.Border = 0;
						testetable.HorizontalAlignment = 0;
						//testetable.PaddingTop = 4f;
						
						PdfPCell testetablepartnumber= new PdfPCell(new Phrase(Itens[19].ToString(),Mixed));//Colunas									
						testetablepartnumber.Border = 0;
						testetablepartnumber.HorizontalAlignment = 0;
						//testetable.PaddingTop = 4f;
						
						PdfPCell testetablecaixa = new PdfPCell(new Phrase(Itens[21].ToString(),Mixed));//Colunas									
						testetablecaixa.Border = 0;
						testetablecaixa.HorizontalAlignment = 0;
						//testetablecaixa.PaddingTop = 4f;
												
					  	
//						PdfPCell ORD1 = new PdfPCell(new Phrase(Itens[0].ToString(),Mixed));//Colunas									
//						ORD1.Border = 0;
//						ORD1.HorizontalAlignment = 0;
						
						
						
						
						string textoItens ="";
						//List<string> linhas;						
						linhas = SeparaLinhas(Itens[28].ToString(),50);	
						//linhas.FindIndex();
						for (int i = 0; i < linhas.Count; i++) {
							textoItens = textoItens + linhas[i].ToString() + "\n";
						}
						
						
						
						
						Frete =	Convert.ToDouble(String.Format("{0:0.0000}", Itens[29]));
						TotalAMT = Itens[32].ToString();
						
						
						PdfPCell PROD1 = new PdfPCell(new Phrase(textoItens.ToString(),Mixed));//Colunas
						PROD1.Border = 0;
						PROD1.HorizontalAlignment = 0;
						PROD1.Colspan =7;
						//PROD1.PaddingTop = 4f;
						//PROD1.PaddingLeft =4f;
						PROD1.PaddingBottom=4f;
						
						
												
						PdfPCell IDENT1 = new PdfPCell(new Phrase(Itens[37].ToString()	,Mixed));//Colunas						/
						IDENT1.Border = 0;
						IDENT1.HorizontalAlignment = 1;			           				
						
						PdfPCell CO1 = new PdfPCell(new Phrase(Itens[22].ToString(),Mixed));//Colunas					
						CO1.Border = 0;
						CO1.HorizontalAlignment = 0;						
			
						PdfPCell UM1 = new PdfPCell(new Phrase(Itens[23].ToString(),Mixed));//Colunas						
						UM1.Border = 0;
						UM1.HorizontalAlignment = 0;								
						
						PdfPCell QTY1 = new PdfPCell(new Phrase(Itens[24].ToString(),Mixed));//Colunas									
						QTY1.Border = 0;
						QTY1.HorizontalAlignment =0;
						
						string unitcost = Itens[25].ToString();
						//unitcost.ToString("F3", CultureInfo.InvariantCulture);
			
						PdfPCell COST1 = new PdfPCell(new Phrase(unitcost,Mixed));//Colunas
						COST1.Border = 0;
						COST1.HorizontalAlignment = 1;
						
						 						
			
						PdfPCell AMT1 = new PdfPCell(new Phrase(Itens[26].ToString(),Mixed));//Colunas									
						AMT1.Border = 0;
						AMT1.HorizontalAlignment = 1;
						
						PdfPCell CTYPE1 = new PdfPCell(new Phrase(Itens[27].ToString(),font8));//Colunas									
						CTYPE1.Border = 0;
						CTYPE1.HorizontalAlignment = 1;
						Ctype =Itens[27].ToString();
						
					
					
						table2.AddCell(testetableCASE);			
						table2.AddCell(testetable);											
						table2.AddCell(testetablevazia);			
						table2.AddCell(testetablevazia);
						table2.AddCell(testetablevazia);			
						table2.AddCell(testetablevazia);
						table2.AddCell(testetablevazia);
						table2.AddCell(testetablevazia);
						table2.AddCell(testetablevazia);
						
						
						
						table3.AddCell(testetablepartnumber);			
						table3.AddCell(testetablecaixa);											
						table3.AddCell(IDENT1);			
						table3.AddCell(CO1);
						table3.AddCell(UM1);			
						table3.AddCell(QTY1);
						table3.AddCell(COST1);
						table3.AddCell(AMT1);
						table3.AddCell(CTYPE1);
						
						table.AddCell(testetablevazia);
						table.AddCell(PROD1);			
//						table.AddCell(PROD1);											
//						table.AddCell(IDENT1);			
//						table.AddCell(CO1);
//						table.AddCell(UM1);			
//						table.AddCell(QTY1);
//						table.AddCell(COST1);
//						table.AddCell(AMT1);
//						table.AddCell(CTYPE1);

						//MessageBox.Show(table2.Rows.Count.ToString());
						//MessageBox.Show(table3.Rows.Count.ToString());
						contador =	table2.Rows.Count +	table3.Rows.Count+1;				
						
					//	MessageBox.Show(contador.ToString());
						
											
						if (yesprimeira)
						{
							
							
							if (contadorLinhas ==30)
							{
								SetCabecalho(documento,1,writer,rs);
							}						

							if(contadorLinhas - linhas.Count >= 0)
							{
								
								contadorLinhas =contadorLinhas -linhas.Count-contador;
								ImprimeItens(writer,documento,table2);
								ImprimeItens(writer,documento,table3);
								ImprimeItens(writer,documento,table);								
								table.DeleteBodyRows();
								table2.DeleteBodyRows();
								table3.DeleteBodyRows();
								 
								contadorItens =contadorItens +contador+linhas.Count;
//								if (contador+linhas.Count < 5)
//								{
//									ImprimeCaixasyesPrimeira(documento,writer);
//									
//								}	
								
							}
							
							else
							{
								SetCabecalho(documento,2,writer,rs);
								contadorLinhas = 42;
								contadorLinhas =contadorLinhas -linhas.Count-contador;
								ImprimeItens(writer,documento,table2);
								ImprimeItens(writer,documento,table3);
								ImprimeItens(writer,documento,table);
								
								table.DeleteBodyRows();
								table2.DeleteBodyRows();
								table3.DeleteBodyRows();								
								yesprimeira =false;								
							}							
								
							
						}
						else
						{
//							if(contadorLinhas == 2)
//									contadorLinhas = 2;
							if(contadorLinhas - linhas.Count >= 0)
							{
								contadorLinhas =contadorLinhas -linhas.Count-contador;
								ImprimeItens(writer,documento,table2);
								ImprimeItens(writer,documento,table3);
								ImprimeItens(writer,documento,table);
								
								//contadorLinhas =contadorLinhas -linhas.Count;
								table.DeleteBodyRows();	
								table2.DeleteBodyRows();
								table3.DeleteBodyRows();								
								
							}
							else
							{
								SetCabecalho(documento,2,writer,rs);
								ImprimeItens(writer,documento,table2);
								ImprimeItens(writer,documento,table3);
								ImprimeItens(writer,documento,table);								
								table.DeleteBodyRows();
								table2.DeleteBodyRows();
								table3.DeleteBodyRows();								
								contadorLinhas = 42;
								contadorLinhas =contadorLinhas -linhas.Count-contador;
								
									
							}
							
						}							
						terms = Itens[33].ToString();
						dterms =Itens[14].ToString();
						unit = Itens[30].ToString();
		}
				
		
					
		
			int pageN = writer.PageNumber;
			int Total =pageN;           
			for (int i = 0; i < pageN; i++) 
			{
				Total++;
				
			}
			if (Total > Total -1)
			{
				
			
				
				rodape(documento,TotalAMT,Frete,Ctype,dterms,unit,Total);
				if(contadorItens < 5)
				{
					Phrase phrase = null;					
					BaseFont HelveticaBold = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, false);
					Font fontHelveticaBold = new Font(HelveticaBold, 10,0 ,GrayColor.BLACK);				
					BaseFont HelveticaBold2 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, false);
					Font fontHelveticaBold2 = new Font(HelveticaBold, 12,0 ,GrayColor.BLACK);
		
					ImprimeCaixasyesPrimeira(documento,writer);
					documento.NewPage();
					LogoDocumento(documento,writer);
					tabelaIvoiceTipoCab2(documento,writer, rs);
					
					string quebra = "  ";				
					documento.Add(new Paragraph(quebra)); //pula uma linha
					documento.Add(new Paragraph(quebra));
					documento.Add(new Paragraph(quebra));
					documento.Add(new Paragraph(quebra));
					
					PdfPCell logo = null;
					PdfPCell teste = new PdfPCell(new Phrase("UNIT OF MEASURE:\nTERMS OF PAYMENT:\n ",fontHelveticaBold));//Colunas			
					teste.UseVariableBorders = true;
					teste.BorderColorRight = BaseColor.WHITE;
					teste.HorizontalAlignment = 0;					
					teste.Colspan =1;	
					teste.PaddingTop = 0f;
		       		teste.Border =0;	       		
				
		            table3 = new PdfPTable(2);
		            table3.TotalWidth = 580f;
		            table3.LockedWidth = true;
		            table3.SetWidths(new float[] { 0.1f, 0.4f});

              
					table3.AddCell(teste); 					
		            phrase = new Phrase();	            
		            phrase.Add(new Chunk(unit+"\n", FontFactory.GetFont("Helvetica", 10)));
		            phrase.Add(new Chunk(terms, FontFactory.GetFont("Helvetica", 10)));           
		            logo = PhraseCell(phrase, 0); 
					table3.AddCell(logo); 

					Chunk chunk2 = new Chunk("CERTIFIED TRUE AND CORRECT",fontHelveticaBold);
		            Paragraph p2 = new Paragraph();
		            p2.Alignment = Element.ALIGN_LEFT;
		            p2.Add(chunk2);
		            
		            Chunk chunk3 = new Chunk("IBM BRASIL LTDA\nLAST PAGE",fontHelveticaBold);
		            Paragraph p3 = new Paragraph();
		            p3.Alignment = Element.ALIGN_RIGHT;
		            p3.Add(chunk3);
		            
					Paragraph ph = new Paragraph("");
					ph.Add(new Chunk("\n"));	            
		           
		            table3.AddCell(logo);          
		        	documento.Add(table3);
		        	documento.Add(ph);
		        	documento.Add(ph);		        	
					documento.Add(p2);
					documento.Add(p3);       					            
					
					
				}
				else{
									
				documento.NewPage();					
				LogoDocumento(documento,writer);
				tabelaIvoiceTipoCab2(documento,writer, rs);	
				ImprimeCaixas(documento,writer);
				}
							
			}
		return table; 
	}
		
		public void ImprimeItens(PdfWriter writer,Document documento, PdfPTable table)
		{			 
			documento.Add(table);
			
					
		}
		
		private void SetColunns(PdfWriter writer,Document documento,IBM rs)
		{
			
			BaseFont Helvetica = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, false);			
			Font fontHelvetica = new Font(Helvetica, 10, 2 ,GrayColor.BLACK);
			Font Mixed = FontFactory.GetFont("Mixed", 9);
			
			BaseFont HelveticaBold = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, false);
			Font fontHelveticaBold = new Font(HelveticaBold, 10,0 ,GrayColor.BLACK);				
			BaseFont HelveticaBold2 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, false);
			Font fontHelveticaBold2 = new Font(HelveticaBold, 12,0 ,GrayColor.BLACK);
			var table = new PdfPTable(9) { WidthPercentage = 100}; var colWidthPercentages2 = new[] { 10f, 30f, 10f, 10f,10f,10f,15f,15f,10f }; table.SetWidths(colWidthPercentages2);
			var table2 = new PdfPTable(9) { WidthPercentage = 100}; var colWidthPercentages3 = new[] { 10f, 30f, 10f, 10f,10f,10f,15f,15f,10f }; table2.SetWidths(colWidthPercentages3);
			table2.DefaultCell.Border = 0;
			
			
			PdfPCell testetablevazia = new PdfPCell(new Phrase("    ",Mixed));//Colunas									
			testetablevazia.Border = 0;
			testetablevazia.HorizontalAlignment =0;
			
			PdfPCell testetable = new PdfPCell(new Phrase(rs.Ident.ToString(),Mixed));//Colunas
			testetable.Border = 0;
			testetable.HorizontalAlignment = 1;
			PdfContentByte cb = writer.DirectContent;			

			cb.SetLineWidth(0.5f);

            cb.Rectangle(6, 415, 585,30);
  			cb.Stroke();
            // Don't forget to call the BeginText() method when done doing graphics!
            cb.BeginText();           
            
            cb.EndText();

	
			PdfPCell ORD = new PdfPCell(new Phrase("\nORD",fontHelveticaBold));//Colunas
			//cellApelido.BackgroundColor = new BaseColor(10, 200, 10);
			ORD.UseVariableBorders = true;
			ORD.BorderColorRight = BaseColor.WHITE;
			ORD.HorizontalAlignment = 0;			
			ORD.PaddingBottom = 10f;			
			ORD.Colspan =1;	
			ORD.Border =0;				
			
			
			PdfPCell PROD = new PdfPCell(new Phrase("\nPROD\nDESCRIPTION",fontHelveticaBold));//Colunas
			//cellNome.BackgroundColor = new BaseColor(10, 200, 10);
			PROD.UseVariableBorders = true;
			PROD.HorizontalAlignment =0;			
			PROD.Colspan =1;		
			PROD.BorderColorRight = BaseColor.WHITE;
			PROD.BorderColorLeft = BaseColor.WHITE;
			PROD.Border =0;	
			
			
			
			PdfPCell IDENT = new PdfPCell(new Phrase("\nIDENT",fontHelveticaBold));//Colunas
			//cellEmail.BackgroundColor = new BaseColor(10, 200, 10);
			IDENT.UseVariableBorders = true;
			IDENT.HorizontalAlignment = 0;			
			IDENT.Colspan =1;
			IDENT.BorderColorRight = BaseColor.WHITE;
			IDENT.BorderColorLeft = BaseColor.WHITE;
			IDENT.Border =0;
						
						
			
			PdfPCell CO = new PdfPCell(new Phrase("\nC/O",fontHelveticaBold));//Colunas	
			//cellSenha.BackgroundColor = new BaseColor(10, 200, 10);			
			CO.Colspan = 1;
			CO.HorizontalAlignment = 0;			
			CO.UseVariableBorders = true;
			CO.BorderColorRight = BaseColor.WHITE;
			CO.BorderColorLeft = BaseColor.WHITE;
			CO.Border =0;

			PdfPCell UM = new PdfPCell(new Phrase("\nUM",fontHelveticaBold));//Colunas
			//cellTipo.BackgroundColor = new BaseColor(10, 200, 10);
			UM.HorizontalAlignment = 0;			
			UM.Colspan =1;	
			UM.UseVariableBorders = true;
			UM.BorderColorRight = BaseColor.WHITE;
			UM.BorderColorLeft = BaseColor.WHITE;			
			UM.Border =0;		
			
			PdfPCell QTY = new PdfPCell(new Phrase("\nQTY",fontHelveticaBold));//Colunas
			//cellDataSenha.BackgroundColor = new BaseColor(204, 204, 204); // Color Cinza(204, 204, 204)			
			QTY.Colspan = 1;
			QTY.HorizontalAlignment = 0;
			QTY.UseVariableBorders = true;
			QTY.BorderColorRight = BaseColor.WHITE;
			QTY.BorderColorLeft = BaseColor.WHITE;			
			QTY.Border =0;			

			PdfPCell COST = new PdfPCell(new Phrase("\nUNI COST",fontHelveticaBold));//Colunas
			//cellDataSenha.BackgroundColor = new BaseColor(204, 204, 204); // Color Cinza(204, 204, 204)			
			COST.Colspan = 1;
			COST.HorizontalAlignment = 0;				
			COST.UseVariableBorders = true;
			COST.BorderColorRight = BaseColor.WHITE;
			COST.BorderColorLeft = BaseColor.WHITE;	
			COST.Border =0;			
						

			PdfPCell AMT = new PdfPCell(new Phrase("\nAMT($USA)",fontHelveticaBold));//Colunas
			//cellDataSenha.BackgroundColor = new BaseColor(204, 204, 204); // Color Cinza(204, 204, 204)			
			AMT.Colspan = 1;
			AMT.HorizontalAlignment =0;
			AMT.UseVariableBorders = true;
			AMT.BorderColorRight = BaseColor.WHITE;
			AMT.BorderColorLeft = BaseColor.WHITE;
			AMT.Border =0;
			
			
			PdfPCell CTYPE = new PdfPCell(new Phrase("\nCTYPE",fontHelveticaBold));//Colunas
			//cellDataSenha.BackgroundColor = new BaseColor(204, 204, 204); // Color Cinza(204, 204, 204)	
			CTYPE.UseVariableBorders = true;			
			CTYPE.Colspan = 1;
			CTYPE.HorizontalAlignment =0;
			CTYPE.BorderColorLeft = BaseColor.WHITE;
			CTYPE.Border =0;			

			table.AddCell(ORD);			
			table.AddCell(PROD);			
			table.AddCell(IDENT);			
			table.AddCell(CO);
			table.AddCell(UM);			
			table.AddCell(QTY);
			table.AddCell(COST);			
			table.AddCell(AMT);
			table.AddCell(CTYPE);
			
			table2.AddCell(testetablevazia);			
			table2.AddCell(testetablevazia);			
			table2.AddCell(testetable);			
			table2.AddCell(testetablevazia);
			table2.AddCell(testetablevazia);			
			table2.AddCell(testetablevazia);
			table2.AddCell(testetablevazia);			
			table2.AddCell(testetablevazia);
			table2.AddCell(testetablevazia);
			
			documento.Add(table);
			documento.Add(table2);
		
		}
		
		
		
		private void SetColunns2(PdfWriter writer,Document documento,IBM rs)
		{
			
			BaseFont HelveticaBold = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, false);
			Font fontHelveticaBold = new Font(HelveticaBold, 10,0 ,GrayColor.BLACK);				
			BaseFont HelveticaBold2 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, false);
			Font fontHelveticaBold2 = new Font(HelveticaBold, 12,0 ,GrayColor.BLACK);
			var table = new PdfPTable(9) { WidthPercentage = 100}; var colWidthPercentages2 = new[] { 10f, 30f, 10f, 10f,10f,10f,15f,15f,10f }; table.SetWidths(colWidthPercentages2);
			 PdfContentByte cb = writer.DirectContent;			
					
			cb.SetLineWidth(0.5f);

                    cb.Rectangle(6, 575, 585,30);
          			cb.Stroke();
                    // Don't forget to call the BeginText() method when done doing graphics!
                    cb.BeginText();                   
                    
            cb.EndText();

	
			PdfPCell ORD = new PdfPCell(new Phrase("\nORD",fontHelveticaBold));//Colunas
			//cellApelido.BackgroundColor = new BaseColor(10, 200, 10);
			ORD.UseVariableBorders = true;
			ORD.BorderColorRight = BaseColor.WHITE;
			ORD.HorizontalAlignment = 0;			
			ORD.PaddingBottom = 10f;			
			ORD.Colspan =1;	
			ORD.Border =0;				
			
			
			PdfPCell PROD = new PdfPCell(new Phrase("\nPROD\nDESCRIPTION",fontHelveticaBold));//Colunas
			//cellNome.BackgroundColor = new BaseColor(10, 200, 10);
			PROD.UseVariableBorders = true;
			PROD.HorizontalAlignment =0;			
			PROD.Colspan =1;		
			PROD.BorderColorRight = BaseColor.WHITE;
			PROD.BorderColorLeft = BaseColor.WHITE;
			PROD.Border =0;	
			
			PdfPCell IDENT = new PdfPCell(new Phrase("\nIDENT",fontHelveticaBold));//Colunas
			//cellEmail.BackgroundColor = new BaseColor(10, 200, 10);
			IDENT.UseVariableBorders = true;
			IDENT.HorizontalAlignment = 0;			
			IDENT.Colspan =1;
			IDENT.BorderColorRight = BaseColor.WHITE;
			IDENT.BorderColorLeft = BaseColor.WHITE;
			IDENT.Border =0;	
						
			
			PdfPCell CO = new PdfPCell(new Phrase("\nC/O",fontHelveticaBold));//Colunas	
			//cellSenha.BackgroundColor = new BaseColor(10, 200, 10);			
			CO.Colspan = 1;
			CO.HorizontalAlignment = 0;			
			CO.UseVariableBorders = true;
			CO.BorderColorRight = BaseColor.WHITE;
			CO.BorderColorLeft = BaseColor.WHITE;
			CO.Border =0;

			PdfPCell UM = new PdfPCell(new Phrase("\nUM",fontHelveticaBold));//Colunas
			//cellTipo.BackgroundColor = new BaseColor(10, 200, 10);
			UM.HorizontalAlignment = 0;			
			UM.Colspan =1;	
			UM.UseVariableBorders = true;
			UM.BorderColorRight = BaseColor.WHITE;
			UM.BorderColorLeft = BaseColor.WHITE;			
			UM.Border =0;		
			
			PdfPCell QTY = new PdfPCell(new Phrase("\nQTY",fontHelveticaBold));//Colunas
			//cellDataSenha.BackgroundColor = new BaseColor(204, 204, 204); // Color Cinza(204, 204, 204)			
			QTY.Colspan = 1;
			QTY.HorizontalAlignment = 0;
			QTY.UseVariableBorders = true;
			QTY.BorderColorRight = BaseColor.WHITE;
			QTY.BorderColorLeft = BaseColor.WHITE;			
			QTY.Border =0;			

			PdfPCell COST = new PdfPCell(new Phrase("\nUNI COST",fontHelveticaBold));//Colunas
			//cellDataSenha.BackgroundColor = new BaseColor(204, 204, 204); // Color Cinza(204, 204, 204)			
			COST.Colspan = 1;
			COST.HorizontalAlignment = 0;				
			COST.UseVariableBorders = true;
			COST.BorderColorRight = BaseColor.WHITE;
			COST.BorderColorLeft = BaseColor.WHITE;	
			COST.Border =0;			
						

			PdfPCell AMT = new PdfPCell(new Phrase("\nAMT($USA)",fontHelveticaBold));//Colunas
			//cellDataSenha.BackgroundColor = new BaseColor(204, 204, 204); // Color Cinza(204, 204, 204)			
			AMT.Colspan = 1;
			AMT.HorizontalAlignment =0;
			AMT.UseVariableBorders = true;
			AMT.BorderColorRight = BaseColor.WHITE;
			AMT.BorderColorLeft = BaseColor.WHITE;
			AMT.Border =0;
			
			
			PdfPCell CTYPE = new PdfPCell(new Phrase("\nCTYPE",fontHelveticaBold));//Colunas
			//cellDataSenha.BackgroundColor = new BaseColor(204, 204, 204); // Color Cinza(204, 204, 204)	
			CTYPE.UseVariableBorders = true;			
			CTYPE.Colspan = 1;
			CTYPE.HorizontalAlignment =0;
			CTYPE.BorderColorLeft = BaseColor.WHITE;
			CTYPE.Border =0;			

			table.AddCell(ORD);			
			table.AddCell(PROD);			
			table.AddCell(IDENT);			
			table.AddCell(CO);
			table.AddCell(UM);			
			table.AddCell(QTY);
			table.AddCell(COST);			
			table.AddCell(AMT);
			table.AddCell(CTYPE);
			documento.Add(table);
		
		}
		
		 private void writeText(PdfContentByte cb, string Text, int X, int Y, BaseFont font, int Size)
        {
            cb.SetFontAndSize(font, Size);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Text, X, Y, 0);
        }
		
		private static PdfPCell PhraseCell(Phrase phrase, int align)
	    {
	        PdfPCell cell = new PdfPCell(phrase);
	        //cell.BorderColor = Color.WHITE;        				
				cell.Colspan =1;	
				cell.PaddingTop = 0f;
	       // cell.PaddingBottom = 2f;
	        
	       cell.Border =0;		           	
	       
	        return cell;
	    }
	    private static PdfPCell ImageCell(string path, float scale, int align)
	    {
	        iTextSharp.text.Image image = iTextSharp.text.Image.GetInstance(path);
	        image.ScalePercent(scale);
	        PdfPCell cell = new PdfPCell(image);
	        cell.Border =0;
	       // cell.VerticalAlignment = PdfPCell.ALIGN_TOP;
	       // cell.HorizontalAlignment = align;
	        cell.PaddingBottom = 0f;
	        cell.PaddingTop =35f;
	        return cell;
	    }
		     
	     private void LogoDocumento(Document documento,PdfWriter writer)
	     {
	     	Phrase phrase = null;
			PdfPCell logo = null;
			PdfPTable table3 = null;
			string quebra = "  ";
			BaseFont HelveticaBold = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, false);
			Font fontHelveticaBold = new Font(HelveticaBold, 10,0 ,GrayColor.BLACK);				
			BaseFont HelveticaBold2 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, false);
			Font fontHelveticaBold2 = new Font(HelveticaBold, 12,0 ,GrayColor.BLACK);
			
			documento.Add(new Paragraph(quebra)); //pula uma linha
			documento.Add(new Paragraph(quebra));
			documento.Add(new Paragraph(quebra));
			documento.Add(new Paragraph(quebra));
			
			PdfPCell teste = new PdfPCell(new Phrase("                   FROM:\n      HEAD OFFICE:\n  SUMARE PLANT:\n                  C.G.C.",fontHelveticaBold));//Colunas			
			teste.UseVariableBorders = true;
			teste.BorderColorRight = BaseColor.WHITE;
			teste.HorizontalAlignment = 2;					
			teste.Colspan =1;	
			teste.PaddingTop = 0f;
       		teste.Border =0;	       		
		
            table3 = new PdfPTable(3);
            table3.TotalWidth = 580f;
            table3.LockedWidth = true;
            table3.SetWidths(new float[] { 0.1f, 0.2f, 0.7f });

            //Company Logo
            logo = ImageCell(@"\\saobr3k401\d$\MULTAPP\Aplicativos\Debug\LogoIBM.png", 40f, PdfPCell.ALIGN_CENTER);
            table3.AddCell(logo);  
			table3.AddCell(teste); 
			
            phrase = new Phrase();	            
            phrase.Add(new Chunk("IBM BRASIL - INDUSTIA, MAQUINAS E SERVIÇOS LTDA\n", FontFactory.GetFont("Mixed", 10)));
            phrase.Add(new Chunk("AV. PASTEUR, 138/146-22290-RIO JANEIRO\n", FontFactory.GetFont("Mixed", 10)));
            phrase.Add(new Chunk("ROD. SP. 101 -TR. CAMPINAS MONTE MOR, KM 9 - C.P.71 \n", FontFactory.GetFont("Mixed",10)));
            phrase.Add(new Chunk("33.372.251/0062-78   INSC EST 748.000.503.112", FontFactory.GetFont("Mixed", 10)));
            logo = PhraseCell(phrase, 0); 



			           
			
			Chunk chunk2 = new Chunk("INVOICE\n",fontHelveticaBold2);
            Paragraph p2 = new Paragraph();
            p2.Alignment = Element.ALIGN_CENTER;
            p2.Add(chunk2);
            
			Paragraph ph = new Paragraph("");
			ph.Add(new Chunk("\n"));	            
           
            table3.AddCell(logo);          
        	documento.Add(table3);
        	
			documento.Add(p2);
			documento.Add(ph);
	     }
	     
	     private void tabelaIvoiceTipoCab1(Document documento, PdfWriter writer,IBM rs)
	     {
			
	     	
					
						
	     	BaseFont HelveticaBold = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, false);
	     	BaseFont HelveticaNormal = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, false);
	     	Font HelveticaN =new Font(HelveticaNormal, 10,0 ,GrayColor.BLACK);
			Font fontHelveticaBold = new Font(HelveticaBold, 10,0 ,GrayColor.BLACK);				
			BaseFont HelveticaBold2 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, false);
			Font fontHelveticaBold2 = new Font(HelveticaBold, 12,0 ,GrayColor.BLACK);				
			Font Mixed = FontFactory.GetFont("Mixed", 9);
			//BaseFont f_cb = BaseFont.CreateFont("c:\\windows\\fonts\\calibrib.ttf", BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
			
			Font Mixed2 = FontFactory.GetFont("Mixed", 10);		
			
			PdfContentByte cb = writer.DirectContent;
			
			cb.BeginText();		  
			cb.SetFontAndSize(BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, false), 10);
			cb.SetTextMatrix(8, 700);
			cb.ShowText(rs.NroArquivo);
			cb.EndText();
			
			cb.BeginText();		  
			cb.SetFontAndSize(BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, false), 10);
			cb.SetTextMatrix(40, 680);
			cb.ShowText("INVOICE TO\n");
			cb.EndText();
			
			int lastwriteposition = 100;
			int top_margin = 660;
			int left_margin = 8;
			
			
				
			
			PdfContentByte cb8 = writer.DirectContent;
			cb8.BeginText();
			
			
			List<string> linhas;						
			linhas = SeparaLinhas(rs.invTO.ToString(),34);
			for (int i = 0; i < linhas.Count; i++)
                    {
                        
                        writeText(cb8, linhas[i].ToString(), left_margin, top_margin, HelveticaNormal, 10);                      
                        
                        top_margin -= 12;

                        // Implement a page break function, checking if the write position has reached the lastwriteposition
                        if(top_margin <= lastwriteposition)
                        {
                            // We need to end the writing before we change the page
                            cb8.EndText();
                            // Make the page break
                           // document.NewPage();
                            // Start the writing again
                            cb8.BeginText();
                            // Assign the new write location on page two!
                            // Here you might want to implement a new header function for the new page
                            top_margin = 660;
                        }
                    }
			
			
				cb8.EndText();			

			// INVOICE TO -------------------------------FIM------------------------------
			
			
			// SHIP TO -------------------------------------------------------------------
			
			
			PdfContentByte cb2 = writer.DirectContent;
			cb2.BeginText();		  
			cb2.SetFontAndSize(BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, false), 10);
			cb2.SetTextMatrix(40, 570);
			cb2.ShowText("SHIP TO");
			cb2.EndText();
			
			int lastwritepositionShip = 100;
			int top_marginShip = 550;
			int left_marginShip = 8;
			 
			
			PdfContentByte cb14 = writer.DirectContent;
			cb14.BeginText();  
			
			
			List<string> linhasShip;						
			linhasShip = SeparaLinhas(rs.ShipTO.ToString(),34);
			for (int x = 0; x < linhasShip.Count; x++)
                    {
                        
                        writeText(cb14, linhasShip[x].ToString(), left_marginShip, top_marginShip, HelveticaNormal, 10);                      
                        
                        top_marginShip -= 12;

                        // Implement a page break function, checking if the write position has reached the lastwriteposition
                        if(top_marginShip <= lastwritepositionShip)
                        {
                            // We need to end the writing before we change the page
                            cb14.EndText();
                            // Make the page break
                           // document.NewPage();
                            // Start the writing again
                            cb14.BeginText();
                            // Assign the new write location on page two!
                            // Here you might want to implement a new header function for the new page
                            top_marginShip = 550;
                        }
                    }
			
			
			
			cb14.EndText();

			// SHIP TO ------------------------------FIM-------------------------------------	

			PdfContentByte cb3 = writer.DirectContent;
			cb3.BeginText();		  
			cb3.SetFontAndSize(BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, false), 10);
			cb3.SetTextMatrix(8, 460);
			cb3.ShowText("METHOD OF TRANSPORT:");
			cb3.EndText();
		
			PdfContentByte cb7 = writer.DirectContent;
			cb7.BeginText();		  
			cb7.SetFontAndSize(BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, false), 10);
			cb7.SetTextMatrix(8, 450);
			cb7.ShowText(rs.Transp);
			cb7.EndText();
			
			PdfContentByte cb4 = writer.DirectContent;
			cb4.BeginText();		  
			cb4.SetFontAndSize(BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, false), 10);
			cb4.SetTextMatrix(300, 470);
			cb4.ShowText("DELIVERY TERMS:");
			cb4.EndText();
		  
			PdfContentByte cb6 = writer.DirectContent;
			cb6.BeginText();		  
			cb6.SetFontAndSize(BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, false), 10);
			cb6.SetTextMatrix(300, 460);
			cb6.ShowText(rs.DTerms);
			cb6.EndText();			
			
			PdfPTable table1 = new PdfPTable(3);
			table1.DefaultCell.Border = 0;
			table1.WidthPercentage = 50;
			table1.HorizontalAlignment = Element.ALIGN_RIGHT;

			

			PdfPCell cell11 = new PdfPCell();
			cell11.AddElement(new Paragraph("        INVOICE DATE\n"+"        DD MM YYYY",fontHelveticaBold));//Colunas
			//cellEmail.BackgroundColor = new BaseColor(10, 200, 10);
			cell11.HorizontalAlignment = 1;			
			cell11.Colspan =1;  
			cell11.AddElement(new Phrase("           "+rs.DataFatura+"",Mixed));				
			
			PdfPCell cell12 = new PdfPCell();
			cell12.AddElement(new Paragraph("      INVOICE TYPE\n\n", fontHelveticaBold));   
			cell12.AddElement(new Paragraph("        "+rs.TpInvoice+"", Mixed));
			cell12.VerticalAlignment = Element.ALIGN_CENTER;				
			
			PdfPCell cell13 = new PdfPCell();
			cell13.AddElement(new Paragraph("PAGE     "+1+" \n\n", fontHelveticaBold));
			cell13.AddElement(new Paragraph("        "+rs.CaixaCAB+"", Mixed));    
			cell13.VerticalAlignment = Element.ALIGN_CENTER;		     	
	     	
			
			
			table1.AddCell(cell11);
			table1.AddCell(cell12);
			table1.AddCell(cell13);			
			
			PdfPTable table2 = new PdfPTable(4);			
			PdfPCell cell21 = new PdfPCell();
			cell21.AddElement(new Paragraph("       INVOICING CTY / LOC",fontHelveticaBold));
			cell21.Colspan = 2;		
			cell21.HorizontalAlignment =1;
			cell21.AddElement(new Paragraph("                       "+rs.InvCTY+"   "+ rs.InvLoc+"", Mixed));    
			cell21.VerticalAlignment = Element.ALIGN_CENTER;
			
			PdfPCell cell215 = new PdfPCell();
			cell215.AddElement(new Paragraph("       INVOICE TO CTY / LOC",fontHelveticaBold));
			cell215.Colspan = 2;		
			cell215.HorizontalAlignment =1;
			cell215.AddElement(new Paragraph("                       "+rs.InvToCty+"   "+ rs.InvToLoc+"", Mixed));
			cell215.VerticalAlignment = Element.ALIGN_CENTER;				
			
			PdfPCell cell22 = new PdfPCell();
			cell22.AddElement(new Paragraph("       SHIPPING CTY / LOC",fontHelveticaBold));
			cell22.Colspan = 2;		
			cell22.HorizontalAlignment =1;
			cell22.AddElement(new Paragraph("                       "+rs.ShipCTY+"   "+ rs.ShipLoc+"", Mixed));    
			cell22.VerticalAlignment = Element.ALIGN_CENTER;				
			
			PdfPCell cell225 = new PdfPCell();
			cell225.AddElement(new Paragraph("       SHIP TO CTY / LOC",fontHelveticaBold));
			cell225.Colspan = 2;		
			cell225.HorizontalAlignment =1;
			cell225.AddElement(new Paragraph("                       "+rs.ShipToCty+"   "+ rs.ShipToLoc+"", Mixed));    
			cell225.VerticalAlignment = Element.ALIGN_CENTER;				
			
			PdfPCell cell23 = new PdfPCell();
			cell23.AddElement(new Paragraph(""));
			cell23.VerticalAlignment =Element.ALIGN_CENTER;
			cell23.Border = 0;
			
			PdfPCell cell24 = new PdfPCell();
			cell24.AddElement(new Paragraph(""));
			cell24.VerticalAlignment =Element.ALIGN_CENTER;
			cell24.Border = 0;
			
			PdfPCell cellVazia = new PdfPCell();
			cellVazia.AddElement(new Paragraph(""));
			cellVazia.Colspan = 3;
			cellVazia.Border = 0;
			cellVazia.HorizontalAlignment = 0;
			
			PdfPCell cellVazia2 = new PdfPCell();
			cellVazia2.AddElement(new Paragraph(""));
			cellVazia2.VerticalAlignment =Element.ALIGN_CENTER; 	
			cellVazia2.Colspan = 3;
			cellVazia2.Border = 0;
			cellVazia2.HorizontalAlignment = 0; 
			
			PdfPCell cell5 = new PdfPCell(new Phrase("DELIVERY TERMS\n ",fontHelveticaBold));
			//cell5.Border = 0;
			cell5.Colspan = 3;
			cell5.HorizontalAlignment = 0;			
			
			cell5.AddElement(new Paragraph("\n", Mixed)); 
			
			cell5.AddElement(new Paragraph("                    REQ. ORDER # "+rs.PedORder+"\n",Mixed));
			cell5.AddElement(new Paragraph("                       "+rs.ATTN+"\n", Mixed));
			cell5.AddElement(new Paragraph("                       EMERGENCY CODE: "+rs.EMERG+"\n", Mixed));
			cell5.VerticalAlignment =Element.ALIGN_CENTER;
			cell5.HorizontalAlignment = 2; 				
			
			cell5.Border = 0;
			cell5.Colspan = 3;
			cell5.FixedHeight = 130.0f;
			//cell5.HorizontalAlignment = 0;
			
			PdfPCell bottom2 = new PdfPCell(cell5);
			bottom2.Colspan =4;					
			table2.AddCell(cell21);			
			table2.AddCell(cell22);
			table2.AddCell(cell215);			
			table2.AddCell(cell225);
			table2.AddCell(bottom2);
			table2.AddCell(cell23);
			table2.AddCell(cell24);
			table2.AddCell(cellVazia);
			table2.AddCell(cellVazia2);				
			PdfPCell cell2A = new PdfPCell(table2);				
			cell2A.Colspan = 3;				
			table1.AddCell(cell2A);				
			PdfPCell cell41 = new PdfPCell();				
			cell41.AddElement(new Paragraph(""));				
			cell41.AddElement(new Paragraph(""));				
			cell41.VerticalAlignment = Element.ALIGN_LEFT;			
			PdfPCell cell42 = new PdfPCell();				
			cell42.AddElement(new Paragraph(""));				
			cell42.AddElement(new Paragraph(""));			
			cell42.VerticalAlignment = Element.ALIGN_RIGHT;   
			table1.AddCell(cellVazia2);
			
			documento.Add(table1);						

	     }
	     private void tabelaIvoiceTipoCab2(Document documento,PdfWriter writer,IBM rs)
	     {
			BaseFont HelveticaBold = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, false);
			Font fontHelveticaBold = new Font(HelveticaBold, 10,0 ,GrayColor.BLACK);				
			BaseFont HelveticaBold2 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, false);
			Font fontHelveticaBold2 = new Font(HelveticaBold, 12,0 ,GrayColor.BLACK);				
			Font Mixed = FontFactory.GetFont("Mixed", 9);
			
			
			PdfContentByte cb;
            PdfTemplate template;

	 	 	cb = writer.DirectContent;
            template = cb.CreateTemplate(50, 50);

            int pageN = writer.PageNumber;
            String text =  pageN.ToString(); 

			cb.BeginText();		  
			cb.SetFontAndSize(BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, false), 10);
			cb.SetTextMatrix(8, 700);
			cb.ShowText(rs.NroArquivo);
			cb.EndText();            
			
			PdfContentByte cbmethod = writer.DirectContent;
			cbmethod.BeginText();		  
			cbmethod.SetFontAndSize(BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, false), 10);
			cbmethod.SetTextMatrix(9, 680);
			cbmethod.ShowText("   INVOICE TO CTY / LOC" );
			cbmethod.EndText();
			
			PdfContentByte cb7cbmethod = writer.DirectContent;
			cb7cbmethod.BeginText();		  
			cb7cbmethod.SetFontAndSize(BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, false), 10);
			cb7cbmethod.SetTextMatrix(9, 670);
			cb7cbmethod.ShowText("          SHIP TO CTY/ LOC");
			cb7cbmethod.EndText();

		  
			PdfContentByte cbmethod2 = writer.DirectContent;
			cbmethod2.BeginText();		  
			cbmethod2.SetFontAndSize(BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, false), 10);
			cbmethod2.SetTextMatrix(150, 680);
			cbmethod2.ShowText(rs.InvToCty+"   "+rs.InvToLoc );
			cbmethod2.EndText();
			
			
			PdfContentByte cb7cbmethod2 = writer.DirectContent;
			cb7cbmethod2.BeginText();		  
			cb7cbmethod2.SetFontAndSize(BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, false), 10);
			cb7cbmethod2.SetTextMatrix(150, 670);
			cb7cbmethod2.ShowText(rs.ShipToCty+"   "+ rs.ShipToLoc);
			cb7cbmethod2.EndText();
			
			
			PdfContentByte cbtransport = writer.DirectContent;
			cbtransport.BeginText();		  
			cbtransport.SetFontAndSize(BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, false), 10);
			cbtransport.SetTextMatrix(8, 630);
			cbtransport.ShowText("METHOD OF TRANSPORT:" );
			cbtransport.EndText();
			
			PdfContentByte cbtransport2 = writer.DirectContent;
			cbtransport2.BeginText();		  
			cbtransport2.SetFontAndSize(BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, false), 10);
			cbtransport2.SetTextMatrix(8, 620);
			cbtransport2.ShowText(rs.Transp);
			cbtransport2.EndText();
	     	
	     	PdfPCell cell11 = new PdfPCell();
			cell11.AddElement(new Paragraph("        INVOICE DATE\n"+"        DD MM YYYY",fontHelveticaBold));//Colunas
			//cellEmail.BackgroundColor = new BaseColor(10, 200, 10);
			cell11.HorizontalAlignment = 1;			
			cell11.Colspan =1;  
			cell11.AddElement(new Phrase("           "+rs.DataFatura+"",Mixed));				
			
			PdfPCell cell12 = new PdfPCell();
			cell12.AddElement(new Paragraph("      INVOICE TYPE\n\n", fontHelveticaBold));   
			cell12.AddElement(new Paragraph("        "+rs.TpInvoice+"", Mixed));
			cell12.VerticalAlignment = Element.ALIGN_CENTER;				
			
			PdfPCell cell13 = new PdfPCell();
			cell13.AddElement(new Paragraph("PAGE     "+text+" \n\n", fontHelveticaBold));
			cell13.AddElement(new Paragraph("        "+rs.CaixaCAB+"", Mixed));    
			cell13.VerticalAlignment = Element.ALIGN_CENTER;		     	
	     	
	     	PdfPTable tableDocumento23 = new PdfPTable(3);
		    tableDocumento23.WidthPercentage = 50;
		    tableDocumento23.HorizontalAlignment = Element.ALIGN_RIGHT;       
		            
		    PdfPCell cell233 = new PdfPCell();
			cell233.AddElement(new Paragraph("DELIVERY TERMS:", fontHelveticaBold));	
		    cell233.AddElement(new Paragraph(rs.DTerms,Mixed));   
			cell233.HorizontalAlignment = 0;    
		    cell233.Colspan = 3;
		    
		    tableDocumento23.AddCell(cell11);
		    tableDocumento23.AddCell(cell12);
		    tableDocumento23.AddCell(cell13);
		    tableDocumento23.AddCell(cell233);		   
		    
		    PdfPTable table40 = new PdfPTable(1);
		    table40.WidthPercentage = 100;
		    
		    PdfPCell linhacelula = new PdfPCell(new Phrase("",fontHelveticaBold));//Colunas	
			//cellSenha.BackgroundColor = new BaseColor(10, 200, 10);			
			linhacelula.Colspan = 1;
			linhacelula.HorizontalAlignment = 0;			
			linhacelula.UseVariableBorders = true;			
			linhacelula.BorderColorBottom = BaseColor.WHITE;
			linhacelula.BorderColorRight = BaseColor.WHITE;
			linhacelula.BorderColorLeft = BaseColor.WHITE;			            
		    table40.AddCell(linhacelula);
       		documento.Add(tableDocumento23);
			documento.Add(table40);
			
	     	
	     }
	     private void rodape(Document documento, string valuegoods,double frete,string ctype,string dterms, string unit,int total)
	     {
	     	BaseFont HelveticaBold = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, false);
			Font fontHelveticaBold = new Font(HelveticaBold, 10,0 ,GrayColor.BLACK);				
			BaseFont HelveticaBold2 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, false);
			Font fontHelveticaBold2 = new Font(HelveticaBold, 12,0 ,GrayColor.BLACK);
			BaseFont Helvetica = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, false);			
			Font Mixed = FontFactory.GetFont("Mixed", 9);
			Font fontHelvetica = new Font(Helvetica, 10, 2 ,GrayColor.BLACK);
			
	     	Chunk c1= new Chunk("_____________\n", fontHelvetica);
		    Chunk c2 = new Chunk("VALUE OF GOODS  ", fontHelveticaBold);		   
		    Chunk c3 = new Chunk("              "+valuegoods+"\n",Mixed);
		    Chunk c4= new Chunk("_____________\n", fontHelvetica);
		    Chunk c5 = new Chunk("TOTAL FOB", fontHelveticaBold);		   
		    Chunk c6 = new Chunk("              "+valuegoods+" "+ctype+"\n",Mixed);
		    Chunk c7 = new Chunk("INTERNATIONAL FREIGHT", fontHelveticaBold);		   
		    Chunk c8 = new Chunk("              "+frete+" "+ctype+"\n",Mixed);
		    Chunk c9= new Chunk("_____________\n", fontHelvetica);
		    Chunk c91= new Chunk(dterms, Mixed);
		    Chunk c92 = new Chunk("              "+(Convert.ToDouble(valuegoods)+frete)+"\n",Mixed);
		    Chunk c111= new Chunk("UNIT OF MEASURE:  ", fontHelveticaBold);
		    Chunk cunit= new Chunk(unit, Mixed);
		   
		    
		    Paragraph p = new Paragraph();		   
			p.Alignment = Element.ALIGN_RIGHT;
    		p.Add(c1);
		    p.Add(c2);
		    p.Add(c3);
			p.Add(c4);
		    p.Add(c5);
		    p.Add(c6);
			p.Add(c7);
			p.Add(c8);
			p.Add(c9);
			p.Add(c91);
			p.Add(c92);
			
			documento.Add(p);
			
			if (total !=2)
			{
			Paragraph p3 = new Paragraph();
			p3.Alignment = Element.ALIGN_LEFT;
			p3.Add(c111);
			p3.Add(cunit);
			documento.Add(p3);	
			}
			
			
							
			
	     }
	     
	    
	 	 public void OnEndPage(PdfWriter writer, Document document)
        {
           PdfContentByte cb;
           PdfTemplate template;

	 	 	cb = writer.DirectContent;
            template = cb.CreateTemplate(50, 50);

            int pageN = writer.PageNumber;
            String text = "Page " + pageN.ToString();
           

            iTextSharp.text.Rectangle pageSize = document.PageSize;

            cb.SetRGBColorFill(100, 100, 100);

            cb.BeginText();
            
          //  cb.SetFontAndSize(this.RunDateFont.BaseFont, this.RunDateFont.Size);
            cb.SetTextMatrix(document.LeftMargin, pageSize.GetBottom(document.BottomMargin));
            cb.ShowText(text);

            cb.EndText();

            cb.AddTemplate(template, document.LeftMargin , pageSize.GetBottom(document.BottomMargin));
            //return text;
        }
	 	 
	 	 private void ImprimeCaixas(Document documento,PdfWriter writer)
	 	 {
	 	 	double TotalGross =	0;				
			double TotalNET = 0;
			double TotalCubic =0;
	 	 	string quebra = "  ";
	 	 	string casecode="";
	 	 	string origin="";
	 	 	string terms ="";
	 	 	string qtdcaixa="";
			BaseFont HelveticaBold = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, false);
			Font fontHelveticaBold = new Font(HelveticaBold, 10,0 ,GrayColor.BLACK);				
			BaseFont HelveticaBold2 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, false);
			Font fontHelveticaBold2 = new Font(HelveticaBold, 12,0 ,GrayColor.BLACK);
			Font Mixed = FontFactory.GetFont("Mixed", 9);
	 	 	
	 	 	documento.Add(new Paragraph(quebra));
			documento.Add(new Paragraph(quebra));
			documento.Add(new Paragraph(quebra));
// Tabela Documento3

			
				
			var tableDocumento3 = new PdfPTable(10) { WidthPercentage = 100}; var colWidthPercentages3 = new[] { 15f, 10f, 10f, 10f,15f,10f,10f,15f,10f,10f }; tableDocumento3.SetWidths(colWidthPercentages3);
			
			

			PdfPCell CASE = new PdfPCell(new Phrase("CASE  NUMBER",fontHelveticaBold));//Colunas
			//cellApelido.BackgroundColor = new BaseColor(10, 200, 10);
			CASE.UseVariableBorders = true;
			CASE.BorderColorRight = BaseColor.WHITE;
			CASE.HorizontalAlignment = 0;						
			CASE.Colspan =1;	
			CASE.Border =0;				
			
			PdfPCell CODE = new PdfPCell(new Phrase("CASE CODE",fontHelveticaBold));//Colunas
			//cellNome.BackgroundColor = new BaseColor(10, 200, 10);
			CODE.UseVariableBorders = true;
			CODE.HorizontalAlignment =0;			
			CODE.Colspan =1;		
			CODE.BorderColorRight = BaseColor.WHITE;
			CODE.BorderColorLeft = BaseColor.WHITE;
			CODE.Border =0;	
			
			PdfPCell EACH = new PdfPCell(new Phrase("EACH         L",fontHelveticaBold));//Colunas
			//cellEmail.BackgroundColor = new BaseColor(10, 200, 10);
			EACH.UseVariableBorders = true;
			EACH.HorizontalAlignment = 1;			
			EACH.Colspan =1;
			EACH.BorderColorRight = BaseColor.WHITE;
			EACH.BorderColorLeft = BaseColor.WHITE;			
			EACH.Border =0;			
			
			PdfPCell X = new PdfPCell(new Phrase(" ",fontHelveticaBold));//Colunas	
			//cellSenha.BackgroundColor = new BaseColor(10, 200, 10);			
			X.Colspan = 1;
			X.HorizontalAlignment = 1;			
			X.UseVariableBorders = true;
			X.BorderColorRight = BaseColor.WHITE;
			X.BorderColorLeft = BaseColor.WHITE;
			X.Border =0;	

			PdfPCell DIMENSION = new PdfPCell(new Phrase("DIMENSION          W",fontHelveticaBold));//Colunas
			//cellTipo.BackgroundColor = new BaseColor(10, 200, 10);
			DIMENSION.HorizontalAlignment = 1;			
			DIMENSION.Colspan =1;	
			DIMENSION.UseVariableBorders = true;
			DIMENSION.BorderColorRight = BaseColor.WHITE;
			DIMENSION.BorderColorLeft = BaseColor.WHITE;			
			DIMENSION.Border =0;		
			
			PdfPCell X2 = new PdfPCell(new Phrase(" ",fontHelveticaBold));//Colunas
			//cellDataSenha.BackgroundColor = new BaseColor(204, 204, 204); // Color Cinza(204, 204, 204)			
			X2.Colspan = 1;
			X2.HorizontalAlignment = 1;
			X2.UseVariableBorders = true;
			X2.BorderColorRight = BaseColor.WHITE;
			X2.BorderColorLeft = BaseColor.WHITE;			
			X2.Border =0;			

			PdfPCell CM = new PdfPCell(new Phrase("(CM)           H",fontHelveticaBold));//Colunas
			//cellDataSenha.BackgroundColor = new BaseColor(204, 204, 204); // Color Cinza(204, 204, 204)			
			CM.Colspan = 1;
			CM.HorizontalAlignment = 1;				
			CM.UseVariableBorders = true;
			CM.BorderColorRight = BaseColor.WHITE;
			CM.BorderColorLeft = BaseColor.WHITE;
			CM.Border =0;					

			PdfPCell WEIGHT = new PdfPCell(new Phrase("WEIGHT GROSS",fontHelveticaBold));//Colunas
			//cellDataSenha.BackgroundColor = new BaseColor(204, 204, 204); // Color Cinza(204, 204, 204)			
			WEIGHT.Colspan = 1;
			WEIGHT.HorizontalAlignment =0;
			WEIGHT.UseVariableBorders = true;
			WEIGHT.BorderColorRight = BaseColor.WHITE;
			WEIGHT.BorderColorLeft = BaseColor.WHITE;
			WEIGHT.Border =0;
			
			PdfPCell KILOS = new PdfPCell(new Phrase("KILOS   NET",fontHelveticaBold));//Colunas
			//cellDataSenha.BackgroundColor = new BaseColor(204, 204, 204); // Color Cinza(204, 204, 204)	
			KILOS.UseVariableBorders = true;			
			KILOS.Colspan = 1;
			KILOS.HorizontalAlignment =0;
			KILOS.BorderColorLeft = BaseColor.WHITE;
			KILOS.BorderColorRight = BaseColor.WHITE;
			KILOS.Border =0;			

			PdfPCell CUBIC = new PdfPCell(new Phrase("CUBIC METER",fontHelveticaBold));//Colunas
			//cellDataSenha.BackgroundColor = new BaseColor(204, 204, 204); // Color Cinza(204, 204, 204)	
			CUBIC.UseVariableBorders = true;			
			CUBIC.Colspan = 1;
			CUBIC.HorizontalAlignment =0;
			CUBIC.BorderColorLeft = BaseColor.WHITE;
			CUBIC.Border =0;
			
			tableDocumento3.AddCell(CASE);			
			tableDocumento3.AddCell(CODE);			
			tableDocumento3.AddCell(EACH);			
			tableDocumento3.AddCell(X);
			tableDocumento3.AddCell(DIMENSION);			
			tableDocumento3.AddCell(X2);
			tableDocumento3.AddCell(CM);			
			tableDocumento3.AddCell(WEIGHT);
			tableDocumento3.AddCell(KILOS);
			tableDocumento3.AddCell(CUBIC);
			
		//Celulas para os dados
			SqlConnection conn;
			conn = dbConecta();	
			string query = " m_IBMExpRelCaixa '" + Ano + "', '" + Processo + "', '" + Invoice + "'";
			
			conn = dbConecta();				
		     	
		
			 SqlCommand cmd = new SqlCommand(query, conn);
			
			  try
			
			  {
			
			    conn.Open();
			
			    using (SqlDataReader rdr = cmd.ExecuteReader())
			
			    {
			
			      while (rdr.Read())
			
			      { 
		
				PdfPCell CASEcell = new PdfPCell(new Phrase(rdr[14].ToString(),Mixed));//Colunas
				//cellApelido.BackgroundColor = new BaseColor(10, 200, 10);
				CASEcell.UseVariableBorders = true;
				CASEcell.BorderColorRight = BaseColor.WHITE;
				CASEcell.HorizontalAlignment = 0;						
				CASEcell.Colspan =1;	
				CASEcell.Border =0;		
				
				PdfPCell CODEcell = new PdfPCell(new Phrase(rdr[20].ToString(),Mixed));//Colunas
				//cellNome.BackgroundColor = new BaseColor(10, 200, 10);
				CODEcell.UseVariableBorders = true;
				CODEcell.HorizontalAlignment =0;			
				CODEcell.Colspan =1;		
				CODEcell.BorderColorRight = BaseColor.WHITE;
				CODEcell.BorderColorLeft = BaseColor.WHITE;
				CODEcell.Border =0;	
				double eac =Convert.ToDouble(rdr[17]);
				PdfPCell EACcell = new PdfPCell(new Phrase(eac.ToString("F", CultureInfo.InvariantCulture),Mixed));//Colunas
				//cellEmail.BackgroundColor = new BaseColor(10, 200, 10);
				EACcell.UseVariableBorders = true;
				EACcell.HorizontalAlignment = 1;			
				EACcell.Colspan =1;
				EACcell.BorderColorRight = BaseColor.WHITE;
				EACcell.BorderColorLeft = BaseColor.WHITE;			
				EACcell.Border =0;			
				
				PdfPCell Xcell = new PdfPCell(new Phrase("X",fontHelveticaBold));//Colunas	
				//cellSenha.BackgroundColor = new BaseColor(10, 200, 10);			
				Xcell.Colspan = 1;
				Xcell.HorizontalAlignment = 1;			
				Xcell.UseVariableBorders = true;
				Xcell.BorderColorRight = BaseColor.WHITE;
				Xcell.BorderColorLeft = BaseColor.WHITE;
				Xcell.Border =0;	
				double DIMENSION1 =Convert.ToDouble(rdr[16]);
				PdfPCell DIMENSIONcell = new PdfPCell(new Phrase(DIMENSION1.ToString("F", CultureInfo.InvariantCulture),Mixed));//Colunas
				//cellTipo.BackgroundColor = new BaseColor(10, 200, 10);
				DIMENSIONcell.HorizontalAlignment = 1;			
				DIMENSIONcell.Colspan =1;	
				DIMENSIONcell.UseVariableBorders = true;
				DIMENSIONcell.BorderColorRight = BaseColor.WHITE;
				DIMENSIONcell.BorderColorLeft = BaseColor.WHITE;			
				DIMENSIONcell.Border =0;		
				
				PdfPCell X2cell = new PdfPCell(new Phrase("X",fontHelveticaBold));//Colunas
				//cellDataSenha.BackgroundColor = new BaseColor(204, 204, 204); // Color Cinza(204, 204, 204)			
				X2cell.Colspan = 1;
				X2cell.HorizontalAlignment = 1;
				X2cell.UseVariableBorders = true;
				X2cell.BorderColorRight = BaseColor.WHITE;
				X2cell.BorderColorLeft = BaseColor.WHITE;			
				X2cell.Border =0;			
				double CM1 =Convert.ToDouble(rdr[15]);
				PdfPCell CMcell = new PdfPCell(new Phrase(CM1.ToString("F", CultureInfo.InvariantCulture),Mixed));//Colunas
				//cellDataSenha.BackgroundColor = new BaseColor(204, 204, 204); // Color Cinza(204, 204, 204)			
				CMcell.Colspan = 1;
				CMcell.HorizontalAlignment = 1;				
				CMcell.UseVariableBorders = true;
				CMcell.BorderColorRight = BaseColor.WHITE;
				CMcell.BorderColorLeft = BaseColor.WHITE;
				CMcell.Border =0;					
				double  WEIGHT1 =Convert.ToDouble(rdr[18]);
				PdfPCell WEIGHTcell = new PdfPCell(new Phrase(WEIGHT1.ToString("F", CultureInfo.InvariantCulture),Mixed));//Colunas
				//cellDataSenha.BackgroundColor = new BaseColor(204, 204, 204); // Color Cinza(204, 204, 204)			
				WEIGHTcell.Colspan = 1;
				WEIGHTcell.HorizontalAlignment =0;
				WEIGHTcell.UseVariableBorders = true;
				WEIGHTcell.BorderColorRight = BaseColor.WHITE;
				WEIGHTcell.BorderColorLeft = BaseColor.WHITE;
				WEIGHTcell.Border =0;
				double  KILO =Convert.ToDouble(rdr[19]);
				PdfPCell KILOcell = new PdfPCell(new Phrase(KILO.ToString("F", CultureInfo.InvariantCulture),Mixed));//Colunas
				//cellDataSenha.BackgroundColor = new BaseColor(204, 204, 204); // Color Cinza(204, 204, 204)	
				KILOcell.UseVariableBorders = true;			
				KILOcell.Colspan = 1;
				KILOcell.HorizontalAlignment =0;
				KILOcell.BorderColorLeft = BaseColor.WHITE;
				KILOcell.BorderColorRight = BaseColor.WHITE;
				KILOcell.Border =0;		
				double  CUBIC1 =Convert.ToDouble(rdr[21]);
				PdfPCell CUBICcell = new PdfPCell(new Phrase(CUBIC1.ToString("F3", CultureInfo.InvariantCulture),Mixed));//Colunas
				//cellDataSenha.BackgroundColor = new BaseColor(204, 204, 204); // Color Cinza(204, 204, 204)	
				CUBICcell.UseVariableBorders = true;			
				CUBICcell.Colspan = 1;
				CUBICcell.HorizontalAlignment =0;
				CUBICcell.BorderColorLeft = BaseColor.WHITE;
				CUBICcell.Border =0;
				
				qtdcaixa =rdr[22].ToString();
				casecode=rdr[23].ToString();
				origin=rdr[24].ToString();
				terms=rdr[25].ToString();
				
				
				TotalGross =Convert.ToDouble(rdr[26]);			
				TotalNET =Convert.ToDouble(rdr[27]);				
				TotalCubic =Convert.ToDouble(rdr[28]);

				tableDocumento3.AddCell(CASEcell);			
				tableDocumento3.AddCell(CODEcell);			
				tableDocumento3.AddCell(EACcell);			
				tableDocumento3.AddCell(Xcell);
				tableDocumento3.AddCell(DIMENSIONcell);			
				tableDocumento3.AddCell(X2cell);
				tableDocumento3.AddCell(CMcell);			
				tableDocumento3.AddCell(WEIGHTcell);
				tableDocumento3.AddCell(KILOcell);
				tableDocumento3.AddCell(CUBICcell);

			      }
			
			    }
			
			  }
			
			  catch(Exception ex)
			
			  {
			
			    MessageBox.Show(ex.Message);
			
			  }
			  
			conn.Close();
			
//			string textoItens ="";
//			List<string> linhas;
//						//List<string> linhas;						
//						linhas = SeparaLinhas(origin.ToString(),40);	
//						//linhas.FindIndex();
//						for (int i = 0; i < linhas.Count; i++) {
//							textoItens = textoItens +"     " + linhas[i].ToString() + "\n";
//							
//					}
//						
			
			Chunk c111= new Chunk("UNIT OF MEASURE: ", fontHelveticaBold);
		    Chunk c222 = new Chunk("001 - PIECE EACH  ", Mixed);
			Chunk c1113= new Chunk("CASE CODE: ", fontHelveticaBold);
			Chunk c1114= new Chunk(casecode, Mixed);			
			//Chunk c11134= new Chunk("ORIGIM:", fontHelveticaBold);
			//Chunk c111346= new Chunk("     " + textoItens.ToString() + "\n", Mixed);
			//Chunk c111346= new Chunk(origin, Mixed);
			Chunk c11135= new Chunk("TERMS OF PAYMENT: ", fontHelveticaBold);
			Chunk c111356= new Chunk(terms, Mixed);
			Chunk c11136= new Chunk("CERTIFIED TRUE AND CORRET ", fontHelveticaBold);	
			Chunk c11137= new Chunk("IBM BRASIL LTDA"+"\n"+"LAST PAGE", Mixed);																		
			Chunk c12= new Chunk("___________  ________  __________  ", fontHelveticaBold);
			Chunk c13= new Chunk("TOTAL NUMBER OF CASE: ",fontHelveticaBold);                                                                                                 
			Chunk c14= new Chunk(qtdcaixa+"                                                                                                             "+TotalGross.ToString("F", CultureInfo.InvariantCulture)+"                    "+TotalNET.ToString("F", CultureInfo.InvariantCulture)+"            "+TotalCubic.ToString("F3", CultureInfo.InvariantCulture)+"  ", Mixed);
			
			PdfContentByte Linha = writer.DirectContent;
			Linha.BeginText();		  
			Linha.SetFontAndSize(BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, false), 10);
			Linha.SetTextMatrix(8, 533);
			Linha.ShowText("____________  _________  _________________________________________________  ____________________  __________");
			Linha.EndText();
		
			
			var tableOrigin = new PdfPTable(2) { WidthPercentage = 30}; var colWidthPercentages4 = new[] { 7f, 15f }; tableOrigin.SetWidths(colWidthPercentages4);
			tableOrigin.HorizontalAlignment = Element.ALIGN_LEFT;
			PdfPCell Origem = new PdfPCell(new Phrase("ORIGIN:",fontHelveticaBold));//Colunas
				//cellApelido.BackgroundColor = new BaseColor(10, 200, 10);
				Origem.UseVariableBorders = true;
				Origem.BorderColorRight = BaseColor.WHITE;
				Origem.HorizontalAlignment =2;						
				Origem.Colspan =1;	
				Origem.Border =0;
				Origem.PaddingBottom=10f;				
				
				PdfPCell dadosOrgim = new PdfPCell(new Phrase(origin.ToString(),Mixed));//Colunas
				//cellNome.BackgroundColor = new BaseColor(10, 200, 10);
				dadosOrgim.UseVariableBorders = true;
				dadosOrgim.HorizontalAlignment =0;			
				dadosOrgim.Colspan =1;		
				dadosOrgim.BorderColorRight = BaseColor.WHITE;
				dadosOrgim.BorderColorLeft = BaseColor.WHITE;
				dadosOrgim.Border =0;
				dadosOrgim.PaddingBottom=10f;				
				
				tableOrigin.AddCell(Origem);			
				tableOrigin.AddCell(dadosOrgim);
			
			
			Paragraph p4 = new Paragraph();
			p4.Alignment = Element.ALIGN_RIGHT;
			p4.Add(c12);
		   // p4.Add(c222);	

			Paragraph p7 = new Paragraph();
			p7.Alignment =Element.ALIGN_LEFT;
			p7.Add(c13);
			p7.Add(c14);
		   // p4.Add(c222);	
		   
			Paragraph p9 = new Paragraph();
			p9.Alignment =Element.ALIGN_LEFT;
			p9.Add(c1113);
			p9.Add(c1114);
			
			//Paragraph p10 = new Paragraph();
			//p10.Alignment =Element.ALIGN_MIDDLE;
			//p10.Add(c11134);
			//p10.Add(c111346);
			
			Paragraph p11 = new Paragraph();
			p11.Alignment =Element.ALIGN_LEFT;
			p11.Add(c11135);
			p11.Add(c111356);
			
			Paragraph p12 = new Paragraph();
			p12.Alignment =Element.ALIGN_LEFT;
			p12.Add(c11136);
			
			Paragraph p13 = new Paragraph();
			p13.Alignment = Element.ALIGN_RIGHT;
			p13.Add(c11137);
			
			Paragraph p14 = new Paragraph();
			p14.Alignment = Element.ALIGN_LEFT;
			p14.Add(c111);
			p14.Add(c222);
			
			

          
			documento.Add(tableDocumento3);
			documento.Add(p4);
			documento.Add(p7);
			documento.Add(p9);
			//documento.Add(p10);
			documento.Add(tableOrigin);
			documento.Add(new Paragraph(quebra));
			documento.Add(new Paragraph(quebra));
			documento.Add(p14);
		    documento.Add(p11);
		    documento.Add(new Paragraph(quebra));
			documento.Add(new Paragraph(quebra));			
			documento.Add(p12);
		    documento.Add(p13);
		   
	 	 	
	 	 }	
	 	 
	 	 
	 	  private void ImprimeCaixasyesPrimeira(Document documento,PdfWriter writer)
	 	 {
	 	 	
	 	 	string quebra = "  ";
	 	 	double TotalGross =	0;				
			double TotalNET = 0;
			double TotalCubic =0;
	 	 	string casecode="";
	 	 	string origin="";	 	 	
	 	 	string qtdcaixa="";
			BaseFont HelveticaBold = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, false);
			Font fontHelveticaBold = new Font(HelveticaBold, 10,0 ,GrayColor.BLACK);				
			BaseFont HelveticaBold2 = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, false);
			Font fontHelveticaBold2 = new Font(HelveticaBold, 12,0 ,GrayColor.BLACK);
			Font Mixed = FontFactory.GetFont("Mixed", 9);
	 	 	
	 	 	documento.Add(new Paragraph(quebra));
			documento.Add(new Paragraph(quebra));
			documento.Add(new Paragraph(quebra));
// Tabela Documento3
			var tableDocumento3 = new PdfPTable(10) { WidthPercentage = 100}; var colWidthPercentages3 = new[] { 15f, 10f, 10f, 10f,15f,10f,10f,15f,10f,10f }; tableDocumento3.SetWidths(colWidthPercentages3);
		

			PdfPCell CASE = new PdfPCell(new Phrase("CASE  NUMBER",fontHelveticaBold));//Colunas
			//cellApelido.BackgroundColor = new BaseColor(10, 200, 10);
			CASE.UseVariableBorders = true;
			CASE.BorderColorRight = BaseColor.WHITE;
			CASE.HorizontalAlignment = 0;						
			CASE.Colspan =1;	
			CASE.Border =0;				
			
			PdfPCell CODE = new PdfPCell(new Phrase("CASE CODE",fontHelveticaBold));//Colunas
			//cellNome.BackgroundColor = new BaseColor(10, 200, 10);
			CODE.UseVariableBorders = true;
			CODE.HorizontalAlignment =0;			
			CODE.Colspan =1;		
			CODE.BorderColorRight = BaseColor.WHITE;
			CODE.BorderColorLeft = BaseColor.WHITE;
			CODE.Border =0;	
			
			PdfPCell EACH = new PdfPCell(new Phrase("EACH         L",fontHelveticaBold));//Colunas
			//cellEmail.BackgroundColor = new BaseColor(10, 200, 10);
			EACH.UseVariableBorders = true;
			EACH.HorizontalAlignment = 1;			
			EACH.Colspan =1;
			EACH.BorderColorRight = BaseColor.WHITE;
			EACH.BorderColorLeft = BaseColor.WHITE;			
			EACH.Border =0;			
			
			PdfPCell X = new PdfPCell(new Phrase(" ",fontHelveticaBold));//Colunas	
			//cellSenha.BackgroundColor = new BaseColor(10, 200, 10);			
			X.Colspan = 1;
			X.HorizontalAlignment = 1;			
			X.UseVariableBorders = true;
			X.BorderColorRight = BaseColor.WHITE;
			X.BorderColorLeft = BaseColor.WHITE;
			X.Border =0;	

			PdfPCell DIMENSION = new PdfPCell(new Phrase("DIMENSION          W",fontHelveticaBold));//Colunas
			//cellTipo.BackgroundColor = new BaseColor(10, 200, 10);
			DIMENSION.HorizontalAlignment = 1;			
			DIMENSION.Colspan =1;	
			DIMENSION.UseVariableBorders = true;
			DIMENSION.BorderColorRight = BaseColor.WHITE;
			DIMENSION.BorderColorLeft = BaseColor.WHITE;			
			DIMENSION.Border =0;		
			
			PdfPCell X2 = new PdfPCell(new Phrase(" ",fontHelveticaBold));//Colunas
			//cellDataSenha.BackgroundColor = new BaseColor(204, 204, 204); // Color Cinza(204, 204, 204)			
			X2.Colspan = 1;
			X2.HorizontalAlignment = 1;
			X2.UseVariableBorders = true;
			X2.BorderColorRight = BaseColor.WHITE;
			X2.BorderColorLeft = BaseColor.WHITE;			
			X2.Border =0;			

			PdfPCell CM = new PdfPCell(new Phrase("(CM)           H",fontHelveticaBold));//Colunas
			//cellDataSenha.BackgroundColor = new BaseColor(204, 204, 204); // Color Cinza(204, 204, 204)			
			CM.Colspan = 1;
			CM.HorizontalAlignment = 1;				
			CM.UseVariableBorders = true;
			CM.BorderColorRight = BaseColor.WHITE;
			CM.BorderColorLeft = BaseColor.WHITE;
			CM.Border =0;					

			PdfPCell WEIGHT = new PdfPCell(new Phrase("WEIGHT GROSS",fontHelveticaBold));//Colunas
			//cellDataSenha.BackgroundColor = new BaseColor(204, 204, 204); // Color Cinza(204, 204, 204)			
			WEIGHT.Colspan = 1;
			WEIGHT.HorizontalAlignment =0;
			WEIGHT.UseVariableBorders = true;
			WEIGHT.BorderColorRight = BaseColor.WHITE;
			WEIGHT.BorderColorLeft = BaseColor.WHITE;
			WEIGHT.Border =0;
			
			PdfPCell KILOS = new PdfPCell(new Phrase("KILOS   NET",fontHelveticaBold));//Colunas
			//cellDataSenha.BackgroundColor = new BaseColor(204, 204, 204); // Color Cinza(204, 204, 204)	
			KILOS.UseVariableBorders = true;			
			KILOS.Colspan = 1;
			KILOS.HorizontalAlignment =0;
			KILOS.BorderColorLeft = BaseColor.WHITE;
			KILOS.BorderColorRight = BaseColor.WHITE;
			KILOS.Border =0;			

			PdfPCell CUBIC = new PdfPCell(new Phrase("CUBIC METER",fontHelveticaBold));//Colunas
			//cellDataSenha.BackgroundColor = new BaseColor(204, 204, 204); // Color Cinza(204, 204, 204)	
			CUBIC.UseVariableBorders = true;			
			CUBIC.Colspan = 1;
			CUBIC.HorizontalAlignment =0;
			CUBIC.BorderColorLeft = BaseColor.WHITE;
			CUBIC.Border =0;
			
		
				
			tableDocumento3.AddCell(CASE);			
			tableDocumento3.AddCell(CODE);			
			tableDocumento3.AddCell(EACH);			
			tableDocumento3.AddCell(X);
			tableDocumento3.AddCell(DIMENSION);			
			tableDocumento3.AddCell(X2);
			tableDocumento3.AddCell(CM);			
			tableDocumento3.AddCell(WEIGHT);
			tableDocumento3.AddCell(KILOS);
			tableDocumento3.AddCell(CUBIC);

			//Celulas para os dados
			SqlConnection conn;
			conn = dbConecta();	
			string query = " m_IBMExpRelCaixa '" + Ano + "', '" + Processo + "', '" + Invoice + "'";
			
			conn = dbConecta();				
		     	
		
			 SqlCommand cmd = new SqlCommand(query, conn);
			
			  try
			
			  {
			
			    conn.Open();
			
			    using (SqlDataReader rdr = cmd.ExecuteReader())
			
			    {
			
			      while (rdr.Read())
			
			      { 
		
				PdfPCell CASEcell = new PdfPCell(new Phrase(rdr[14].ToString(),Mixed));//Colunas
				//cellApelido.BackgroundColor = new BaseColor(10, 200, 10);
				CASEcell.UseVariableBorders = true;
				CASEcell.BorderColorRight = BaseColor.WHITE;
				CASEcell.HorizontalAlignment = 0;						
				CASEcell.Colspan =1;	
				CASEcell.Border =0;		
				
				PdfPCell CODEcell = new PdfPCell(new Phrase(rdr[20].ToString(),Mixed));//Colunas
				//cellNome.BackgroundColor = new BaseColor(10, 200, 10);
				CODEcell.UseVariableBorders = true;
				CODEcell.HorizontalAlignment =0;			
				CODEcell.Colspan =1;		
				CODEcell.BorderColorRight = BaseColor.WHITE;
				CODEcell.BorderColorLeft = BaseColor.WHITE;
				CODEcell.Border =0;	
				
				double eac =Convert.ToDouble(rdr[17]);
				PdfPCell EACcell = new PdfPCell(new Phrase(eac.ToString("F", CultureInfo.InvariantCulture),Mixed));//Colunas
				//cellEmail.BackgroundColor = new BaseColor(10, 200, 10);
				EACcell.UseVariableBorders = true;
				EACcell.HorizontalAlignment = 1;			
				EACcell.Colspan =1;
				EACcell.BorderColorRight = BaseColor.WHITE;
				EACcell.BorderColorLeft = BaseColor.WHITE;			
				EACcell.Border =0;			
				
				PdfPCell Xcell = new PdfPCell(new Phrase("X",fontHelveticaBold));//Colunas	
				//cellSenha.BackgroundColor = new BaseColor(10, 200, 10);			
				Xcell.Colspan = 1;
				Xcell.HorizontalAlignment = 1;			
				Xcell.UseVariableBorders = true;
				Xcell.BorderColorRight = BaseColor.WHITE;
				Xcell.BorderColorLeft = BaseColor.WHITE;
				Xcell.Border =0;	
				double DIMENSION1 =Convert.ToDouble(rdr[16]);
				PdfPCell DIMENSIONcell = new PdfPCell(new Phrase(DIMENSION1.ToString("F", CultureInfo.InvariantCulture),Mixed));//Colunas
				//cellTipo.BackgroundColor = new BaseColor(10, 200, 10);
				DIMENSIONcell.HorizontalAlignment = 1;			
				DIMENSIONcell.Colspan =1;	
				DIMENSIONcell.UseVariableBorders = true;
				DIMENSIONcell.BorderColorRight = BaseColor.WHITE;
				DIMENSIONcell.BorderColorLeft = BaseColor.WHITE;			
				DIMENSIONcell.Border =0;		
				
				PdfPCell X2cell = new PdfPCell(new Phrase("X",fontHelveticaBold));//Colunas
				//cellDataSenha.BackgroundColor = new BaseColor(204, 204, 204); // Color Cinza(204, 204, 204)			
				X2cell.Colspan = 1;
				X2cell.HorizontalAlignment = 1;
				X2cell.UseVariableBorders = true;
				X2cell.BorderColorRight = BaseColor.WHITE;
				X2cell.BorderColorLeft = BaseColor.WHITE;			
				X2cell.Border =0;			
				double CM1 =Convert.ToDouble(rdr[15]);
				PdfPCell CMcell = new PdfPCell(new Phrase(CM1.ToString("F", CultureInfo.InvariantCulture),Mixed));//Colunas
				//cellDataSenha.BackgroundColor = new BaseColor(204, 204, 204); // Color Cinza(204, 204, 204)			
				CMcell.Colspan = 1;
				CMcell.HorizontalAlignment = 1;				
				CMcell.UseVariableBorders = true;
				CMcell.BorderColorRight = BaseColor.WHITE;
				CMcell.BorderColorLeft = BaseColor.WHITE;
				CMcell.Border =0;					
				double  WEIGHT1 =Convert.ToDouble(rdr[18]);
				PdfPCell WEIGHTcell = new PdfPCell(new Phrase(WEIGHT1.ToString("F", CultureInfo.InvariantCulture),Mixed));//Colunas
				//cellDataSenha.BackgroundColor = new BaseColor(204, 204, 204); // Color Cinza(204, 204, 204)			
				WEIGHTcell.Colspan = 1;
				WEIGHTcell.HorizontalAlignment =0;
				WEIGHTcell.UseVariableBorders = true;
				WEIGHTcell.BorderColorRight = BaseColor.WHITE;
				WEIGHTcell.BorderColorLeft = BaseColor.WHITE;
				WEIGHTcell.Border =0;
				double  KILO =Convert.ToDouble(rdr[19]);
				PdfPCell KILOcell = new PdfPCell(new Phrase(KILO.ToString("F", CultureInfo.InvariantCulture),Mixed));//Colunas
				//cellDataSenha.BackgroundColor = new BaseColor(204, 204, 204); // Color Cinza(204, 204, 204)	
				KILOcell.UseVariableBorders = true;			
				KILOcell.Colspan = 1;
				KILOcell.HorizontalAlignment =0;
				KILOcell.BorderColorLeft = BaseColor.WHITE;
				KILOcell.BorderColorRight = BaseColor.WHITE;
				KILOcell.Border =0;		
				double  CUBIC1 =Convert.ToDouble(rdr[21]);
				PdfPCell CUBICcell = new PdfPCell(new Phrase(CUBIC1.ToString("F3", CultureInfo.InvariantCulture),Mixed));//Colunas
				//cellDataSenha.BackgroundColor = new BaseColor(204, 204, 204); // Color Cinza(204, 204, 204)	
				CUBICcell.UseVariableBorders = true;			
				CUBICcell.Colspan = 1;
				CUBICcell.HorizontalAlignment =0;
				CUBICcell.BorderColorLeft = BaseColor.WHITE;
				CUBICcell.Border =0;
				
				
				qtdcaixa =rdr[22].ToString();
				casecode=rdr[23].ToString();
				origin=rdr[24].ToString();
				//terms=rdr[25].ToString();	

				TotalGross +=WEIGHT1;
				//TotalGross += TotalGross;
				TotalNET +=KILO;
				//TotalNET +=TotalNET;
				TotalCubic +=CUBIC1;	
				//TotalCubic += TotalCubic;				
			

				tableDocumento3.AddCell(CASEcell);			
				tableDocumento3.AddCell(CODEcell);			
				tableDocumento3.AddCell(EACcell);			
				tableDocumento3.AddCell(Xcell);
				tableDocumento3.AddCell(DIMENSIONcell);			
				tableDocumento3.AddCell(X2cell);
				tableDocumento3.AddCell(CMcell);			
				tableDocumento3.AddCell(WEIGHTcell);
				tableDocumento3.AddCell(KILOcell);
				tableDocumento3.AddCell(CUBICcell);
		
			      }
			
			    }
			
			  }
			
			  catch(Exception ex)
			
			  {
			
			    MessageBox.Show(ex.Message);
			
			  }
			  
			conn.Close();
			
			
			
			
			Chunk c1113= new Chunk("CASE CODE: ", fontHelveticaBold);
			Chunk c1114= new Chunk(casecode, Mixed);			
			Chunk c11134= new Chunk("ORIGIN:", fontHelveticaBold);
			Chunk c111346= new Chunk(origin, Mixed);				
			Chunk c12= new Chunk("___________  ________  __________  ", fontHelveticaBold);
			Chunk c13= new Chunk("TOTAL NUMBER OF CASE: ",fontHelveticaBold);                                                                                                 
			Chunk c14= new Chunk(qtdcaixa+"                                                                                                           "+TotalGross.ToString("F", CultureInfo.InvariantCulture)+"                   "+TotalNET.ToString("F", CultureInfo.InvariantCulture)+"            "+TotalCubic.ToString("F3", CultureInfo.InvariantCulture)+"  ", Mixed);
			
			
			
			PdfContentByte Linha = writer.DirectContent;
			Linha.BeginText();		  
			Linha.SetFontAndSize(BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, false), 10);
			Linha.SetTextMatrix(8, 175);
			Linha.ShowText("____________  _________  _________________________________________________  ____________________  __________");
			Linha.EndText();
		
			
			Paragraph p4 = new Paragraph();
			p4.Alignment = Element.ALIGN_RIGHT;
			p4.Add(c12);
		   // p4.Add(c222);	

			Paragraph p7 = new Paragraph();
			p7.Alignment =Element.ALIGN_LEFT;
			p7.Add(c13);
			p7.Add(c14);
		   // p4.Add(c222);	
		   
			Paragraph p9 = new Paragraph();
			p9.Alignment =Element.ALIGN_LEFT;
			p9.Add(c1113);
			p9.Add(c1114);
			
			Paragraph p10 = new Paragraph();
			p10.Alignment =Element.ALIGN_LEFT;
			p10.Add(c11134);
			p10.Add(c111346);
			
			Paragraph p11 = new Paragraph();
			p11.Alignment =Element.ALIGN_LEFT;			
			
			Paragraph p12 = new Paragraph();
			p12.Alignment =Element.ALIGN_LEFT;
			
			
			Paragraph p13 = new Paragraph();
			p13.Alignment = Element.ALIGN_RIGHT;			

          
			documento.Add(tableDocumento3);
			documento.Add(p4);
			documento.Add(p7);
			documento.Add(p9);
			documento.Add(p10);    
	 	 	
	 	 } 	 
	 	 
	 	 public List<string> SeparaLinhas(string texto, int largura)
		{
			string tx = null;
			int x = 0;
			int lt = 0;
			int l = 0;
			int j = 0;
			int k = 0;
			string C = null;
			List<string> linhas = new List<string>();
			string linha = null;
		
			x = 0;
			tx = texto;
			PROCESSA:
			j = 1;
			lt = tx.Length;
		
			//'retira space, cr e lf superfluos
			while (j < lt) {
				C = tx.Substring(j - 1, 1);
				if ((C == " " | C == "\r" | C ==  "\n")) {
				} else {
					break; // TODO: might not be correct. Was : Exit While
				}
				j = j + 1;
			}
		
			if (j != 1)
				tx = tx.Substring(j - 1, lt - (j - 1));
		
			lt = tx.Length;
			l = lt;
			j = 1;
		
			if (lt < largura) {
				x = 1;
			} else {
				l = largura;
			}
		
			//'procura o primeiro separador
			while (j <= l) {
				C = tx.Substring(j - 1, 1);
				if ((C == " " | C ==  "\r" | C == "\n")) {
				} else {
					j = j + 1;
					continue;
				}
		
				//'Separador é cr ou lf
				if (C != " ") {
					j = j - 1;
					break; // TODO: might not be correct. Was : Exit While
				}
				k = j + 1;
		
				//'procura o proximo separador
				while (k <= l) {
					C = tx.Substring(k - 1, 1);
					if ((C == "\r" | C == "\n")) {
						j = k - 1;
						break; // TODO: might not be correct. Was : Exit While
					}
					if (C == " ")
						j = k;
					k = k + 1;
				}
		
				if (x == 1 & k > l)
					j = l;
				break; // TODO: might not be correct. Was : Exit While
			}
		
			if (j > l)
				j = l;
		
			linha = tx.Substring(0, j);
			linhas.Add(linha);
		
			if (j >= lt) {
			} else {
				tx = tx.Substring(j, (lt - j));
				if (tx != "\r\n")
					goto PROCESSA;
		
			}
		
			return linhas;
		}


            
		
		
  	 
//				 public class IBM
//				{
//					public string Invoice;
//					public string Incoterms;			
//					public string Nome;
//					public string Endereço;
//					public string Cidade;
//					public string Bairro;
//					public string FAS_Calc;
//					public string ExNome;
//					public string ExEndereco;
//					public string ExPAIS;
//					public string Descricao;
//					public string TipoProcesso;
//					public string Data;			
//					public string InvoicingCTY;
//					public string ShippingCTY;
//					public string InvoiceTOCTY;
//					public string ShipToCTY;
//					public string InvoiceType;
//					public string DescricaoInvoice;
//					
//					
//							
//				}
				 
				  public class IBM
				{
					public string DataFatura;
					public string TpInvoice;			
					public string invTO;
					public string ShipTO;
					public string CaixaCAB;
					public string InvCTY;
					public string InvLoc;
					public string ShipCTY;
					public string ShipLoc;
					public string InvToCty;
					public string InvToLoc;
					public string ShipToCty;
					public string ShipToLoc;
					public string Transp;			
					public string DTerms;
					public string PedORder;
					public string ATTN;
					public string EMERG;
					public string Ident;
					public string ORD;
					public string NumeroCaixa;
					public string ITPTN;
					public string CD_PAIS_FABRICANTE;
					public string Unid;
					public string ITQTD;
					public string ITPRU;
					public string ITPRT;
					public string CTYPE;
					public string DescricaoEnglish;
					public string Frete;
					public string Unit;
					public string ITNRO;
					public string NroArquivo;
					public string NomeArq;
					public string Observacoes;
					public string Serial;
					public string Ident2;
				
									
					
							
				}

		
				void Button1Click(object sender, EventArgs e)
				{
					GeraInvoice( Ano,Processo,Invoice);
					
				}
			}	
		
		
		}
	
