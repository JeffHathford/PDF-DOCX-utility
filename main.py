import win32com.client
import wx
import wx.xrc
import wx.dataview
import PyPDF2
from pathlib import Path

def dialog_get_paths(wildcard):
    #app = wx.App(None)
    style = wx.FD_OPEN | wx.FD_FILE_MUST_EXIST | wx.FD_MULTIPLE
    dialog = wx.FileDialog(None, "", wildcard=wildcard, style=style)
    if dialog.ShowModal() == wx.ID_OK:
        path = dialog.GetPaths()
    else:
        path = None
    dialog.Destroy()
    return path

def dialog_get_savepath(wildcard):
    #app = wx.App(None)
    style = wx.FD_SAVE
    dialog = wx.FileDialog(None, "", wildcard=wildcard, style=style)
    if dialog.ShowModal() == wx.ID_OK:
        path = dialog.GetPath()
    else:
        path = None
    dialog.Destroy()
    return path

def dialog_get_dir(message=""):
    #app = wx.App(None)
    dialog = wx.DirDialog (None, message, "", wx.DD_DEFAULT_STYLE)
    if dialog.ShowModal() == wx.ID_OK:
        path = dialog.GetPath()
    else:
        path = None
    dialog.Destroy()
    return path

class MyFrame1 ( wx.Frame ):
	
	src_type = 0		#left notebook page number
	dest_type = 0		#right notebook page number

	src_allow_duplicates = False

	pdf_list = []
	docx_list = []

	src_file_ext = ["*.pdf", "*.docx"]			#use in conjunction with src_type
	src_active_list = [pdf_list, docx_list]		#same

	pdf2doc_list = []
	doc2pdf_list = []
	pdfmerge_list = []

	dest_file_ext = ["*.pdf", "*.docx", "*.pdf"]
	dest_active_list = [pdf2doc_list, doc2pdf_list, pdfmerge_list]


	def __init__( self, parent ):
		wx.Frame.__init__ ( self, parent, id = wx.ID_ANY, title = u"Python Utility - PDF&DOCX", pos = wx.DefaultPosition, size = wx.Size( 928,563 ), style = wx.DEFAULT_FRAME_STYLE|wx.TAB_TRAVERSAL )
		
		self.SetSizeHints( wx.DefaultSize, wx.DefaultSize )
		
		bSizer5 = wx.BoxSizer( wx.VERTICAL )
		
		self.m_panel4 = wx.Panel( self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.TAB_TRAVERSAL )
		bSizer7 = wx.BoxSizer( wx.HORIZONTAL )
		
		bSizer8 = wx.BoxSizer( wx.VERTICAL )
		
		sbSizer2 = wx.StaticBoxSizer( wx.StaticBox( self.m_panel4, wx.ID_ANY, u"File workspace" ), wx.VERTICAL )
		
		self.m_notebook21 = wx.Notebook( sbSizer2.GetStaticBox(), wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, 0 )
		self.m_panel5 = wx.Panel( self.m_notebook21, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.TAB_TRAVERSAL )
		bSizer101 = wx.BoxSizer( wx.VERTICAL )
		
		self.m_dataViewListCtrl1 = wx.dataview.DataViewListCtrl( self.m_panel5, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, 0 )
		bSizer101.Add( self.m_dataViewListCtrl1, 1, wx.ALL|wx.EXPAND, 5 )
		
		
		self.m_panel5.SetSizer( bSizer101 )
		self.m_panel5.Layout()
		bSizer101.Fit( self.m_panel5 )
		self.m_notebook21.AddPage( self.m_panel5, u"PDFs", True )
		self.m_panel6 = wx.Panel( self.m_notebook21, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.TAB_TRAVERSAL )
		bSizer111 = wx.BoxSizer( wx.VERTICAL )
		
		self.m_dataViewListCtrl5 = wx.dataview.DataViewListCtrl( self.m_panel6, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, 0 )
		bSizer111.Add( self.m_dataViewListCtrl5, 1, wx.ALL|wx.EXPAND, 5 )
		
		
		self.m_panel6.SetSizer( bSizer111 )
		self.m_panel6.Layout()
		bSizer111.Fit( self.m_panel6 )
		self.m_notebook21.AddPage( self.m_panel6, u"DOCXs", False )
		
		sbSizer2.Add( self.m_notebook21, 1, wx.EXPAND |wx.ALL, 5 )
		
		self.m_checkBox1 = wx.CheckBox( sbSizer2.GetStaticBox(), wx.ID_ANY, u"Allow loading duplicate files", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.m_checkBox1.SetBackgroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_APPWORKSPACE ) )
		
		sbSizer2.Add( self.m_checkBox1, 0, wx.ALL, 5 )



		bSizer8.Add( sbSizer2, 1, wx.EXPAND, 5 )
		
		bSizer11 = wx.BoxSizer( wx.HORIZONTAL )
		
		
		bSizer11.AddStretchSpacer(1)
		
		self.m_button5 = wx.Button( self.m_panel4, wx.ID_ANY, u"Add Folder", wx.DefaultPosition, wx.DefaultSize, 0 )
		bSizer11.Add( self.m_button5, 3, wx.ALIGN_LEFT|wx.ALL, 5 )
		
		
		bSizer11.AddStretchSpacer(1)
		
		self.m_button6 = wx.Button( self.m_panel4, wx.ID_ANY, u"Add File(s)", wx.DefaultPosition, wx.DefaultSize, 0 )
		bSizer11.Add( self.m_button6, 3, wx.ALL, 5 )
		
		
		bSizer11.AddStretchSpacer(1)
		
		
		bSizer8.Add( bSizer11, 0, wx.EXPAND, 5 )
		
		sbSizer21 = wx.StaticBoxSizer( wx.StaticBox( self.m_panel4, wx.ID_ANY, u"Unload:" ), wx.HORIZONTAL )
		
		self.m_button10 = wx.Button( sbSizer21.GetStaticBox(), wx.ID_ANY, u"PDFs", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.m_button10.SetMaxSize( wx.Size( 60,-1 ) )
		
		sbSizer21.Add( self.m_button10, 3, wx.ALL, 5 )
		
		
		sbSizer21.AddStretchSpacer(1)
		
		self.m_button11 = wx.Button( sbSizer21.GetStaticBox(), wx.ID_ANY, u"DOCXs", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.m_button11.SetMaxSize( wx.Size( 60,-1 ) )
		
		sbSizer21.Add( self.m_button11, 3, wx.ALL, 5 )
		
		
		sbSizer21.AddStretchSpacer(1)
		
		self.m_button12 = wx.Button( sbSizer21.GetStaticBox(), wx.ID_ANY, u"All", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.m_button12.SetMaxSize( wx.Size( 60,-1 ) )
		
		sbSizer21.Add( self.m_button12, 3, wx.ALL, 5 )
		
		
		sbSizer21.AddStretchSpacer(1)
		
		self.m_button13 = wx.Button( sbSizer21.GetStaticBox(), wx.ID_ANY, u"Selected", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.m_button13.SetMaxSize( wx.Size( 60,-1 ) )
		
		sbSizer21.Add( self.m_button13, 3, wx.ALL, 5 )
		
		
		bSizer8.Add( sbSizer21, 0, wx.EXPAND, 5 )
		
		
		bSizer7.Add( bSizer8, 5, wx.EXPAND, 5 )
		
		bSizer9 = wx.BoxSizer( wx.VERTICAL )
		
		
		bSizer9.AddStretchSpacer(1)
		
		self.m_button1 = wx.Button( self.m_panel4, wx.ID_ANY, u">", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.m_button1.SetMaxSize( wx.Size( 40,30 ) )
		
		bSizer9.Add( self.m_button1, 1, wx.ALIGN_CENTER|wx.ALL, 5 )
		
		self.m_button2 = wx.Button( self.m_panel4, wx.ID_ANY, u">>", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.m_button2.SetMaxSize( wx.Size( 40,30 ) )
		
		bSizer9.Add( self.m_button2, 1, wx.ALIGN_CENTER|wx.ALL, 5 )
		
		
		bSizer9.AddStretchSpacer(1)
		
		self.m_button3 = wx.Button( self.m_panel4, wx.ID_ANY, u"<<", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.m_button3.SetMaxSize( wx.Size( 40,30 ) )
		
		bSizer9.Add( self.m_button3, 1, wx.ALIGN_CENTER|wx.ALL, 5 )
		
		self.m_button4 = wx.Button( self.m_panel4, wx.ID_ANY, u"<", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.m_button4.SetMaxSize( wx.Size( 40,30 ) )
		
		bSizer9.Add( self.m_button4, 1, wx.ALIGN_CENTER|wx.ALL, 5 )
		
		
		bSizer9.AddStretchSpacer(1)
		
		
		bSizer7.Add( bSizer9, 1, wx.EXPAND, 5 )
		
		bSizer10 = wx.BoxSizer( wx.VERTICAL )
		
		self.m_notebook2 = wx.Notebook( self.m_panel4, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, 0 )
		self.m_panel7 = wx.Panel( self.m_notebook2, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.TAB_TRAVERSAL )
		bSizer15 = wx.BoxSizer( wx.VERTICAL )
		
		self.m_dataViewListCtrl2 = wx.dataview.DataViewListCtrl( self.m_panel7, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, 0 )
		bSizer15.Add( self.m_dataViewListCtrl2, 1, wx.ALL|wx.EXPAND, 5 )
		
		self.m_button7 = wx.Button( self.m_panel7, wx.ID_ANY, u"Convert All", wx.DefaultPosition, wx.DefaultSize, 0 )
		bSizer15.Add( self.m_button7, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )
		
		
		self.m_panel7.SetSizer( bSizer15 )
		self.m_panel7.Layout()
		bSizer15.Fit( self.m_panel7 )
		self.m_notebook2.AddPage( self.m_panel7, u"PDF to DOCX", True )
		self.m_panel8 = wx.Panel( self.m_notebook2, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.TAB_TRAVERSAL )
		bSizer14 = wx.BoxSizer( wx.VERTICAL )
		
		self.m_dataViewListCtrl21 = wx.dataview.DataViewListCtrl( self.m_panel8, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, 0 )
		bSizer14.Add( self.m_dataViewListCtrl21, 1, wx.ALL|wx.EXPAND, 5 )
		
		self.m_button8 = wx.Button( self.m_panel8, wx.ID_ANY, u"Convert All", wx.DefaultPosition, wx.DefaultSize, 0 )
		bSizer14.Add( self.m_button8, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )
		
		
		self.m_panel8.SetSizer( bSizer14 )
		self.m_panel8.Layout()
		bSizer14.Fit( self.m_panel8 )
		self.m_notebook2.AddPage( self.m_panel8, u"DOCX to PDF", False )
		self.m_panel9 = wx.Panel( self.m_notebook2, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.TAB_TRAVERSAL )
		bSizer13 = wx.BoxSizer( wx.VERTICAL )
		
		self.m_dataViewListCtrl6 = wx.dataview.DataViewListCtrl( self.m_panel9, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, 0 )
		bSizer13.Add( self.m_dataViewListCtrl6, 1, wx.ALL|wx.EXPAND, 5 )
		
		self.m_button9 = wx.Button( self.m_panel9, wx.ID_ANY, u"Merge All", wx.DefaultPosition, wx.DefaultSize, 0 )
		bSizer13.Add( self.m_button9, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )
		
		
		self.m_panel9.SetSizer( bSizer13 )
		self.m_panel9.Layout()
		bSizer13.Fit( self.m_panel9 )
		self.m_notebook2.AddPage( self.m_panel9, u"PDF Merger", False )
		
		bSizer10.Add( self.m_notebook2, 1, wx.EXPAND |wx.ALL, 5 )
		
		self.m_gauge1 = wx.Gauge( self.m_panel4, wx.ID_ANY, 100, wx.DefaultPosition, wx.DefaultSize, wx.GA_HORIZONTAL )
		self.m_gauge1.SetValue( 0 ) 
		bSizer10.Add( self.m_gauge1, 0, wx.ALL|wx.EXPAND, 5 )


		bSizer7.Add( bSizer10, 5, wx.EXPAND, 5 )
		
		
		self.m_panel4.SetSizer( bSizer7 )
		self.m_panel4.Layout()
		bSizer7.Fit( self.m_panel4 )
		bSizer5.Add( self.m_panel4, 1, wx.EXPAND |wx.ALL, 5 )
		
		
		self.SetSizer( bSizer5 )
		self.Layout()
		
		self.Centre( wx.BOTH )

		### data list column appending

		self.m_dataViewListCtrl1.AppendTextColumn("Name")
		self.m_dataViewListCtrl1.AppendTextColumn("Path")
		self.m_dataViewListCtrl1.AppendToggleColumn("")

		self.m_dataViewListCtrl5.AppendTextColumn("Name")
		self.m_dataViewListCtrl5.AppendTextColumn("Path")
		self.m_dataViewListCtrl5.AppendToggleColumn("")


		self.m_dataViewListCtrl2.AppendTextColumn("Name")
		self.m_dataViewListCtrl2.AppendTextColumn("Path")
		self.m_dataViewListCtrl2.AppendToggleColumn("")

		self.m_dataViewListCtrl21.AppendTextColumn("Name")
		self.m_dataViewListCtrl21.AppendTextColumn("Path")
		self.m_dataViewListCtrl21.AppendToggleColumn("")

		self.m_dataViewListCtrl6.AppendTextColumn("#")
		self.m_dataViewListCtrl6.AppendTextColumn("Name")
		self.m_dataViewListCtrl6.AppendTextColumn("Path")
		self.m_dataViewListCtrl6.AppendToggleColumn("")

		### event binding

		self.m_button5.Bind( wx.EVT_BUTTON, self.add_folder )
		self.m_button6.Bind( wx.EVT_BUTTON, self.add_files )

		self.m_button10.Bind( wx.EVT_BUTTON, self.unload_PDFs )
		self.m_button11.Bind( wx.EVT_BUTTON, self.unload_DOCXs )
		self.m_button12.Bind( wx.EVT_BUTTON, self.unload_all )
		self.m_button13.Bind( wx.EVT_BUTTON, self.unload_selected )

		self.m_button1.Bind( wx.EVT_BUTTON, self.item_add )
		self.m_button2.Bind( wx.EVT_BUTTON, self.item_add_all )
		self.m_button3.Bind( wx.EVT_BUTTON, self.item_remove_all )
		self.m_button4.Bind( wx.EVT_BUTTON, self.item_remove )

		self.m_button7.Bind( wx.EVT_BUTTON, self.convert_pdf2doc )
		self.m_button8.Bind( wx.EVT_BUTTON, self.convert_doc2pdf )
		self.m_button9.Bind( wx.EVT_BUTTON, self.merge )

		self.m_notebook21.Bind( wx.EVT_NOTEBOOK_PAGE_CHANGED, self.pageflip_src )
		self.m_notebook2.Bind( wx.EVT_NOTEBOOK_PAGE_CHANGED, self.pageflip_dest )

		self.m_checkBox1.Bind( wx.EVT_CHECKBOX, self.allow_duplicates )

		self.m_dataViewListCtrl1.Bind( wx.dataview.EVT_DATAVIEW_ITEM_VALUE_CHANGED, self.select_src1 )
		self.m_dataViewListCtrl5.Bind( wx.dataview.EVT_DATAVIEW_ITEM_VALUE_CHANGED, self.select_src2 )

		self.m_dataViewListCtrl2.Bind( wx.dataview.EVT_DATAVIEW_ITEM_VALUE_CHANGED, self.select_dest1 )
		self.m_dataViewListCtrl21.Bind( wx.dataview.EVT_DATAVIEW_ITEM_VALUE_CHANGED, self.select_dest2 )
		self.m_dataViewListCtrl6.Bind( wx.dataview.EVT_DATAVIEW_ITEM_VALUE_CHANGED, self.select_dest3 )

	
	def __del__( self ):
		pass

	def fromlist_pdf2docx(self, pdf_list, progressbar):

		#if not Path(out_path).exists():
		#	Path(out_path).mkdir(parents=True, exist_ok=True)

		if len(pdf_list) != 0:	
			progressbar.SetRange(len(pdf_list))
			progressbar.SetValue(0)

			out_path = dialog_get_dir() + "\\"

			if out_path != None:
				word = win32com.client.Dispatch("Word.Application")
				word.visible = 0
				for doc in pdf_list:
					filename = doc.split('\\')[-1]

					in_file = str(Path(doc).resolve())

					wb = word.Documents.Open(in_file)
					out_file = str(Path(out_path + filename[0:-4] + ".docx".format()).resolve())

					wb.SaveAs2(out_file, FileFormat=16) # file format for docx
					wb.Close()

					progressbar.SetValue(progressbar.GetValue() + 1)
				word.Quit()

	def fromlist_docx2pdf(self, docx_list, progressbar):
		#if not Path(out_path).exists():
		#	Path(out_path).mkdir(parents=True, exist_ok=True)

		if len(docx_list) != 0:	
			progressbar.SetRange(len(docx_list))
			progressbar.SetValue(0)

			out_path = dialog_get_dir() + "\\"

			if out_path != None:
				word = win32com.client.Dispatch("Word.Application")
				word.visible = 0

				for doc in docx_list:

					filename = doc.split('\\')[-1]

					in_file = str(Path(doc).resolve())

					wb = word.Documents.Open(in_file)
					out_file = str(Path(out_path + "\\" + filename[0:-5] + ".pdf".format()).resolve())

					wb.SaveAs2(out_file, FileFormat=17) # file format for docx
					wb.Close()

					progressbar.SetValue(progressbar.GetValue() + 1)

				word.Quit()

	def fromlist_pdfmerge(self, pdf_list, progressbar):
		if len(pdf_list) != 0:	
			progressbar.SetRange(len(pdf_list))
			progressbar.SetValue(0)

			merger = PyPDF2.PdfFileMerger()
			for doc in pdf_list:
				merger.append(doc)
				progressbar.SetValue(progressbar.GetValue() + 1)
			out_path = dialog_get_savepath("*.pdf")
			merger.write(out_path)


	def add_folder(self, event):	#m_dataViewListCtrl1 and m_dataViewListCtrl5
		query = [item for item in Path(dialog_get_dir()).iterdir() if Path(item).suffix == self.src_file_ext[self.src_type][1:]]
		if query is not None:
			for path in query:
				shortpath = Path(path).stem
				if [shortpath, str(path), False] not in self.src_active_list[self.src_type] or self.src_allow_duplicates == True:
					if self.src_type == 0:
						self.pdf_list.append([shortpath, str(path), False])
						self.m_dataViewListCtrl1.AppendItem([shortpath, str(path), False])
					else:
						self.docx_list.append([shortpath, str(path), False])
						self.m_dataViewListCtrl5.AppendItem([shortpath, str(path), False])

	def add_files(self, event):

		query = dialog_get_paths(self.src_file_ext[self.src_type])
		if query is not None:
			for path in query:
				shortpath = Path(path).stem
				if [shortpath, str(path), False] not in self.src_active_list[self.src_type] or self.src_allow_duplicates == True:
					if self.src_type == 0:
						self.pdf_list.append([shortpath, str(path), False])
						self.m_dataViewListCtrl1.AppendItem([shortpath, str(path), False])
					else:
						self.docx_list.append([shortpath, str(path), False])
						self.m_dataViewListCtrl5.AppendItem([shortpath, str(path), False])						

	def item_add(self, event):
		if self.src_type == self.dest_type % 2:
			src = self.src_active_list[self.src_type]
			dest = self.dest_active_list[self.dest_type]
			temp = []
			for item in src:
				if item[2] == True:
					temp.append(item)
			for item in temp:
				src.remove(item)
				if self.dest_type == 2:
					num = len(dest)
					dest.append([str(num+1), item[0], item[1], False])
					self.m_dataViewListCtrl6.AppendItem([str(num+1), item[0], item[1], False])
				elif self.dest_type == 1:
					dest.append([item[0], item[1], False])
					self.m_dataViewListCtrl21.AppendItem([item[0], item[1], False])
				else:
					dest.append([item[0], item[1], False])
					self.m_dataViewListCtrl2.AppendItem([item[0], item[1], False])
			
			if self.src_type == 0:
				self.m_dataViewListCtrl1.DeleteAllItems()
				for item in src:
					self.m_dataViewListCtrl1.AppendItem(item)
			else:
				self.m_dataViewListCtrl5.DeleteAllItems()
				for item in src:
					self.m_dataViewListCtrl5.AppendItem(item)
			
	def item_add_all(self, event):
		if self.src_type == 0:

			if self.dest_type == 0:
				for item in self.pdf_list:
					self.m_dataViewListCtrl2.AppendItem([item[0], item[1], False])
					self.pdf2doc_list.append([item[0], item[1], False])
				self.pdf_list.clear()
				self.m_dataViewListCtrl1.DeleteAllItems()

			elif self.dest_type == 2:
				for item in self.pdf_list:
					num = len(self.pdfmerge_list)
					self.m_dataViewListCtrl6.AppendItem([str(num+1), item[0], item[1], False])
					self.pdfmerge_list.append([str(num+1), item[0], item[1], False])
				self.pdf_list.clear()
				self.m_dataViewListCtrl1.DeleteAllItems()

		elif self.src_type == 1 and self.dest_type == 1:
			for item in self.docx_list:
				self.m_dataViewListCtrl21.AppendItem([item[0], item[1], False])
				self.doc2pdf_list.append([item[0], item[1], False])
			self.docx_list.clear()
			self.m_dataViewListCtrl5.DeleteAllItems()


		pass

	def item_remove(self, event):
		if self.src_type == self.dest_type % 2:
			src = self.dest_active_list[self.dest_type]
			dest = self.src_active_list[self.src_type]
			temp = []
			for item in src:
				if self.dest_type == 2:
					if item[3] == True:
						temp.append(item)
				else:
					if item[2] == True:
						temp.append(item)
			for item in temp:
				src.remove(item)
				if self.src_type == 0:
					if self.dest_type == 2:
						dest.append([item[1], item[2], False])
						self.m_dataViewListCtrl1.AppendItem([item[1], item[2], False])
					else:
						dest.append([item[0], item[1], False])
						self.m_dataViewListCtrl1.AppendItem([item[0], item[1], False])
				else:
					if self.dest_type == 2:
						dest.append([item[1], item[2], False])
						self.m_dataViewListCtrl5.AppendItem([item[1], item[2], False])
					else:
						dest.append([item[0], item[1], False])
						self.m_dataViewListCtrl5.AppendItem([item[0], item[1], False])

			if self.dest_type == 0:
				self.m_dataViewListCtrl2.DeleteAllItems()
				for item in src:
					self.m_dataViewListCtrl2.AppendItem(item)
			elif self.dest_type == 1:
				self.m_dataViewListCtrl21.DeleteAllItems()
				for item in src:
					self.m_dataViewListCtrl21.AppendItem(item)
			else:
				self.m_dataViewListCtrl6.DeleteAllItems()
				for index, item in enumerate(src):
					item[0] = str(index+1)
					self.m_dataViewListCtrl6.AppendItem(item)

	def item_remove_all(self, event):
		if self.src_type == 0:

			if self.dest_type == 0:
				for item in self.pdf2doc_list:
					self.m_dataViewListCtrl1.AppendItem([item[0], item[1], False])
					self.pdf_list.append([item[0], item[1], False])
				self.pdf2doc_list.clear()
				self.m_dataViewListCtrl2.DeleteAllItems()

			elif self.dest_type == 2:
				for item in self.pdfmerge_list:
					self.m_dataViewListCtrl1.AppendItem([item[1], item[2], False])
					self.pdf_list.append([item[1], item[2], False])
				self.pdfmerge_list.clear()
				self.m_dataViewListCtrl6.DeleteAllItems()

		elif self.src_type == 1 and self.dest_type == 1:
			for item in self.doc2pdf_list:
				self.m_dataViewListCtrl5.AppendItem([item[0], item[1], False])
				self.docx_list.append([item[0], item[1], False])
			self.doc2pdf_list.clear()
			self.m_dataViewListCtrl21.DeleteAllItems()

	def convert_pdf2doc(self, event):	#m_dataViewListCtrl2
		src_list = [x[1] for x in self.pdf2doc_list]
		self.fromlist_pdf2docx(src_list, self.m_gauge1)
		self.pdf2doc_list.clear()
		self.m_dataViewListCtrl2.DeleteAllItems()

	def convert_doc2pdf(self, event):	#m_dataViewListCtrl21
		src_list = [x[1] for x in self.doc2pdf_list]
		self.fromlist_docx2pdf(src_list, self.m_gauge1)
		self.doc2pdf_list.clear()
		self.m_dataViewListCtrl21.DeleteAllItems()

	def merge(self, event):			#m_dataViewListCtrl6
		src_list = [x[2] for x in self.pdfmerge_list]
		self.fromlist_pdfmerge(src_list, self.m_gauge1)
		self.pdfmerge_list.clear()
		self.m_dataViewListCtrl6.DeleteAllItems()

	def unload_PDFs(self, event):		#m_dataViewListCtrl1
		window_warning = wx.MessageDialog(None, "Are you sure you want to unload all PDFs?", "Warning", wx.ICON_WARNING | wx.YES_NO | wx.NO_DEFAULT)

		if window_warning.ShowModal() == wx.ID_YES:
			self.pdf_list.clear()
			self.m_dataViewListCtrl1.DeleteAllItems()

	def unload_DOCXs(self, event):		#m_dataViewListCtrl5
		window_warning = wx.MessageDialog(None, "Are you sure you want to unload all DOCXs?", "Warning", wx.ICON_WARNING | wx.YES_NO | wx.NO_DEFAULT)

		if window_warning.ShowModal() == wx.ID_YES:
			self.docx_list.clear()
			self.m_dataViewListCtrl5.DeleteAllItems()

	def unload_all(self, event):
		window_warning = wx.MessageDialog(None, "Are you sure you want to unload all files?", "Warning", wx.ICON_WARNING | wx.YES_NO | wx.NO_DEFAULT)

		window_warning.SetExtendedMessage("This action cannot be undone.")

		if window_warning.ShowModal() == wx.ID_YES:
			self.pdf_list.clear()
			self.docx_list.clear()
			self.m_dataViewListCtrl1.DeleteAllItems()
			self.m_dataViewListCtrl5.DeleteAllItems()

	def unload_selected(self, event):
		window_warning = wx.MessageDialog(None, "Are you sure you want to unload selected files?", "Warning", wx.ICON_WARNING | wx.YES_NO | wx.NO_DEFAULT)

		if window_warning.ShowModal() == wx.ID_YES:
			temp_list = []
			for item in self.src_active_list[self.src_type]:
				if item[2] == True:
					temp_list.append(item)
			for item in temp_list:
				self.src_active_list[self.src_type].remove(item)
			if self.src_type == 0:
				self.m_dataViewListCtrl1.DeleteAllItems()
				for item in self.src_active_list[self.src_type]:
					self.m_dataViewListCtrl1.AppendItem(item)
			elif self.src_type == 1:
				self.m_dataViewListCtrl5.DeleteAllItems()
				for item in self.src_active_list[self.src_type]:
					self.m_dataViewListCtrl5.AppendItem(item)

	def pageflip_src(self, event):
		self.src_type = event.GetSelection()

	def pageflip_dest(self, event):
		self.dest_type = event.GetSelection()		#0 and 2 are both PDFs

	def allow_duplicates(self, event):
		self.src_allow_duplicates = (self.src_allow_duplicates + 1) % 2

	def select_src1(self, event):
		row = self.m_dataViewListCtrl1.GetSelectedRow()
		if self.pdf_list[row][2] == True:
			self.pdf_list[row][2] = False
		else:
			self.pdf_list[row][2] = True
	
	def select_src2(self, event):
		row = self.m_dataViewListCtrl5.GetSelectedRow()
		if self.docx_list[row][2] == True:
			self.docx_list[row][2] = False
		else:
			self.docx_list[row][2] = True

	def select_dest1(self, event):
		row = self.m_dataViewListCtrl2.GetSelectedRow()
		if self.pdf2doc_list[row][2] == True:
			self.pdf2doc_list[row][2] = False
		else:
			self.pdf2doc_list[row][2] = True

	def select_dest2(self, event):
		row = self.m_dataViewListCtrl21.GetSelectedRow()
		if self.doc2pdf_list[row][2] == True:
			self.doc2pdf_list[row][2] = False
		else:
			self.doc2pdf_list[row][2] = True

	def select_dest3(self, event):
		row = self.m_dataViewListCtrl6.GetSelectedRow()
		if self.pdfmerge_list[row][3] == True:
			self.pdfmerge_list[row][3] = False
		else:
			self.pdfmerge_list[row][3] = True


def main():

    app = wx.App()
    frame = MyFrame1(None)

    frame.Show()
    app.MainLoop()


if __name__ == '__main__':
    main()