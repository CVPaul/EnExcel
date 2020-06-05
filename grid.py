# -*- coding: utf-8 -*- 

import wx
import wx.xrc
import wx.grid

BOND = 0

class DataDialog(wx.Dialog):
    
    def __init__(self, parent, fields):
        wx.Dialog.__init__ ( self, parent, id = wx.ID_ANY, title = wx.EmptyString, pos = wx.DefaultPosition, size = wx.DefaultSize, style = wx.DEFAULT_DIALOG_STYLE )
        self.SetSizeHints( wx.DefaultSize, wx.DefaultSize )
        self.fields = fields
        bSizer3 = wx.BoxSizer( wx.VERTICAL )
        bSizer4 = wx.BoxSizer( wx.VERTICAL )
        self.m_grid2 = wx.grid.Grid( self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, 0 )
        
        # Grid
        self.m_grid2.CreateGrid(len(self.fields), 1)
        # self.m_grid2.CornerLabelValue = "Test"
        self.m_grid2.SetColLabelValue(0,"值")
        for i in range(len(self.fields)):
            self.m_grid2.SetRowLabelValue(i,self.fields[i])
        self.m_grid2.EnableEditing(True)
        self.m_grid2.EnableGridLines(True)
        self.m_grid2.EnableDragGridSize(False)
        self.m_grid2.SetMargins(0,0)
        
        # Columns
        self.m_grid2.EnableDragColMove(False)
        self.m_grid2.EnableDragColSize(True)
        self.m_grid2.SetColLabelSize(30)
        self.m_grid2.SetColLabelAlignment(wx.ALIGN_CENTRE, wx.ALIGN_CENTRE)
        
        # Rows
        self.m_grid2.EnableDragRowSize(True)
        self.m_grid2.SetRowLabelSize(100)
        self.m_grid2.SetRowLabelAlignment(wx.ALIGN_CENTRE, wx.ALIGN_CENTRE)
        
        # Label Appearance
        
        # Cell Defaults
        self.m_grid2.SetDefaultCellAlignment(wx.ALIGN_LEFT, wx.ALIGN_TOP)
        self.m_grid2.SetDefaultColSize(100)
        bSizer4.Add( self.m_grid2, 0, wx.ALL, BOND)
        
        
        bSizer3.Add( bSizer4, 1, wx.EXPAND, BOND)
        btn_sizer = wx.BoxSizer( wx.HORIZONTAL )
        btn_sizer = self.CreateStdDialogButtonSizer(wx.OK|wx.CANCEL)
        bSizer3.Add(btn_sizer, 0.8, wx.EXPAND, BOND)
        
        self.SetSizer(bSizer3)
        self.Layout()
        bSizer3.Fit(self)
        
        self.Centre( wx.BOTH )
    
    def GetData(self):
        dat = {}
        for i in range(len(self.fields)):
            dat[self.m_grid2.GetRowLabelValue(i)] = self.m_grid2.GetCellValue(i,0)
        return dat