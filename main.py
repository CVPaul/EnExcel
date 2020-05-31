# -*- coding: utf-8 -*- 

import wx
import os
import wx.xrc
import wx.dataview
import pandas as pd

from PIL import Image

TITLE = "EnExcel"
ROOT_NAME = "商品总览"
DATA_PATH = "样例数据.xlsx"

BOND = 0
CACHE_SIZE = 5
IMG_WIN_WIDTH = 640
IMG_WIN_HEIGHT = 480
HEIGHT_SCALE_FAC = 0.04
WIN_DEFAULT_WIDTH = 840
WIN_DEFAULT_HEIGHT = 600

class Canvas(wx.Frame):
    def __init__(self,filename):
        wx.Frame.__init__(self, None, -1, filename, size=(IMG_WIN_WIDTH, IMG_WIN_HEIGHT))
        self.filename=filename
        self.Bind(wx.EVT_SIZE, self.change)
        self.p=wx.Panel(self,-1)
        self.SetBackgroundColour('white')

    def start(self):
        self.p.DestroyChildren()#抹掉原先显示的图片
        self.width, self.height = self.GetSize()
        image = Image.open(self.filename)
        self.x, self.y = image.size
        self.x = self.width / 2 - self.x / 2
        self.y = self.height / 2 - self.y / 2
        self.pic = wx.Image(self.filename, wx.BITMAP_TYPE_ANY).ConvertToBitmap()
        # 通过计算获得图片的存放位置
        self.button = wx.BitmapButton(self.p, -1, self.pic, pos=(self.x, self.y))
        self.p.Fit()

    def change(self, size):#如果检测到框架大小的变化，及时改变图片的位置
        if self.filename != "":
            self.start()
        else:
            pass

class EnExcel (wx.Frame):
    def __init__(self, parent, data):
        wx.Frame.__init__ (self, parent, id = wx.ID_ANY, title = TITLE,
            pos = wx.DefaultPosition, 
            size = wx.Size(WIN_DEFAULT_WIDTH,WIN_DEFAULT_HEIGHT),
            style = wx.DEFAULT_FRAME_STYLE|wx.TAB_TRAVERSAL)
        self.data = data
        self.key_code = 0

        self.SetSizeHints(wx.DefaultSize, wx.DefaultSize)
        
        bSizer5 = wx.BoxSizer(wx.HORIZONTAL)
        
        self.m_splitwin1 = wx.SplitterWindow(self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.SP_3D)
        self.m_splitwin1.Bind(wx.EVT_IDLE, self.m_splitwin1OnIdle)
        
        self.m_panel5 = wx.Panel(self.m_splitwin1, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.TAB_TRAVERSAL)
        bSizer6 = wx.BoxSizer(wx.VERTICAL)
        
        self.m_tree = wx.TreeCtrl(self.m_panel5, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.TR_MULTIPLE|wx.TR_DEFAULT_STYLE)
        self.m_tree.Bind(wx.EVT_TREE_SEL_CHANGING, self.on_tree_change)
        self.m_tree.Bind(wx.EVT_KEY_DOWN,self.on_key_down)
        self.m_tree.Bind(wx.EVT_KEY_UP,self.on_key_up)
        bSizer6.Add(self.m_tree, 0, wx.ALL, 0)
        
        self.m_panel5.SetSizer(bSizer6)
        self.m_panel5.Layout()
        bSizer6.Fit(self.m_panel5)
        self.m_splitwin1.Initialize(self.m_panel5)
        bSizer5.Add(self.m_splitwin1, 1, wx.EXPAND, 0)
        
        self.m_splitwin2 = wx.SplitterWindow(self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.SP_3D)
        self.m_splitwin2.Bind(wx.EVT_IDLE, self.m_splitwin2OnIdle)
        
        self.m_panel6 = wx.Panel(self.m_splitwin2, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.TAB_TRAVERSAL)
        bSizer8 = wx.BoxSizer(wx.VERTICAL)

        self.m_dataVLC = wx.dataview.DataViewListCtrl(self.m_panel6, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, 0)
        for token in self.data.columns:
            self.m_dataVLC.AppendTextColumn(token)
        bSizer8.Add(self.m_dataVLC, 0, wx.ALL, BOND)

        self.m_comboBox = wx.ComboBox(self.m_panel6, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, [], wx.TE_PROCESS_ENTER)
        self.m_comboBox.Bind(wx.EVT_TEXT_ENTER, self.on_search)
        bSizer8.Add(self.m_comboBox, 0, wx.ALL, BOND)
        
        self.m_btnSearch = wx.Button(self.m_panel6, wx.ID_ANY, u"搜索", wx.DefaultPosition, wx.DefaultSize, 0)
        self.m_btnSearch.Bind(wx.EVT_BUTTON,self.on_search) 
        bSizer8.Add(self.m_btnSearch, 0, wx.ALL, BOND)
        
        self.m_panel6.SetSizer(bSizer8)
        self.m_panel6.Layout()
        bSizer8.Fit(self.m_panel6)
        self.m_splitwin2.Initialize(self.m_panel6)
        bSizer5.Add(self.m_splitwin2, 1, wx.EXPAND, BOND)
        
        self.m_splitwin3 = wx.SplitterWindow(self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.SP_3D)
        self.m_splitwin3.Bind(wx.EVT_IDLE, self.m_splitwin3OnIdle)
        
        self.m_panel7 = wx.Panel(self.m_splitwin3, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.TAB_TRAVERSAL)
        bSizer9 = wx.BoxSizer(wx.VERTICAL)

        self.m_btnDetail = wx.Button(self.m_panel7, wx.ID_ANY, u"详细信息", wx.DefaultPosition, wx.DefaultSize, 0)
        self.m_btnDetail.Bind(wx.EVT_BUTTON,self.on_detail) 
        bSizer9.Add(self.m_btnDetail, 0, wx.ALL, BOND)

        self.m_textDetail = wx.StaticText(self.m_panel7, wx.ID_ANY, u"详细信息", wx.DefaultPosition, wx.DefaultSize, 0)
        self.m_textDetail.Wrap(-1)
        bSizer9.Add(self.m_textDetail, 0, wx.ALL, BOND)
        
        self.m_panel7.SetSizer(bSizer9)
        self.m_panel7.Layout()
        bSizer9.Fit(self.m_panel7)
        self.m_splitwin3.Initialize(self.m_panel7)
        bSizer5.Add(self.m_splitwin3, 1, wx.EXPAND, BOND)
        
        self.SetSizer(bSizer5)
        self.Layout()
        self.m_status = self.CreateStatusBar()
        self.m_status.SetFieldsCount(2)
        self.m_status.SetStatusText("名称：",0)
        self.m_status.SetStatusText("数量：",1)
        self.m_menubar5 = wx.MenuBar(0)
        self.m_menu4 = wx.Menu()
        self.m_menubar5.Append(self.m_menu4, u"添加")
        
        self.SetMenuBar(self.m_menubar5)
        self.Centre(wx.BOTH)
    
    def __del__(self):
        pass
    
    def on_key_down(self,event):
        self.key_code = event.GetKeyCode()
    
    def on_key_up(self,event):
        self.key_code = 0

    def on_detail(self,event):
        item = self.m_tree.GetSelections()
        if len(item) == 0:
            path = []
        else:
            path = self.get_tree_select_path(item[-1])
        path = "./"+"/".join(path)
        os.system("open %s"%(path+"/detail.pdf"))
        with open(path+"/describe.txt") as fp:
            self.m_textDetail.SetLabelText(fp.read())
        img = Canvas(path+"/example.png")
        img.start()
        img.Center()
        img.Show()

    def get_tree_select_path(self,item):
        select_path = []
        if not item:
            return select_path
        while True:
            name = self.m_tree.GetItemText(item)
            if name == ROOT_NAME:
                break
            else:
                select_path.append(name)
                item = self.m_tree.GetItemParent(item)
        return select_path[::-1]

    def on_search(self,event):
        query = self.m_comboBox.GetValue()
        score_list = []
        for _, row in self.data.iterrows():
            score_list.append(self.calcSim(query,"|".join(row)))
        dat = self.data.copy()
        dat["score"] = score_list
        dat = dat[dat["score"] > 0].sort_values("score").drop("score",axis=1)
        self.m_status.SetStatusText("搜索：" + query,0)
        if dat.empty:
            self.m_dataVLC.DeleteAllItems()
            self.m_status.SetStatusText("共找到：0条记录",1)
        else:
            self.render_data(dat)
            self.m_status.SetStatusText("共找到：%d条记录"%dat.shape[0],1)
        if self.m_comboBox.GetCount() >= CACHE_SIZE:
            self.m_comboBox.Delete(0)
        self.m_comboBox.Append(query)


    def calcSim(self,query,doc):
        return doc.count(query)

    def on_tree_change(self, event):
        dat_collect = []
        selections = [event.GetItem()]
        if self.key_code in [wx.WXK_CONTROL,wx.WXK_SHIFT]:
            selections += self.m_tree.GetSelections()
        # print("len:",selections,len(self.m_tree.GetSelections()))
        select_info = set()
        for item in selections:
            select_info.add(self.m_tree.GetItemText(item))
            select_path = self.get_tree_select_path(item)
            dat = self.data
            for token in select_path:
                dat = dat.loc[token]
            if isinstance(dat,pd.Series):
                dat = dat.to_frame().T
            dat_collect.append(dat)
        dat = pd.concat(dat_collect).drop_duplicates().sort_values("编号")
        self.render_data(dat)
        self.m_status.SetStatusText("名称："+"+".join(select_info),0)
        self.m_status.SetStatusText("数量：%d"%dat.shape[0],1)

    def render_tree(self):
        self.m_tree.DeleteAllItems()
        root = self.m_tree.AddRoot(ROOT_NAME)
        root1 = self.data.index.get_level_values("分支").unique()
        root2 = {t: self.m_tree.AppendItem(root,t) for t in root1}
        root3, root4 = {}, {}
        for row in self.data.index:
            token = row[1]
            if token not in root3:
                root3[token]=self.m_tree.AppendItem(root2[row[0]],token)
        for row in self.data.index:
            token = row[2]
            # pos = 0
            # filename = "./"+"/".join(row[:pos])+"/detail.pdf"
            # content = "当前节点:%s"%"->".join(row[:pos])
            # content += "\n产品介绍:blabalbal...\n"
            # os.system('cp 需求.pdf %s'%(filename))
            if token not in root4:
                root4[token]=self.m_tree.AppendItem(root3[row[1]],token)
        self.m_tree.ExpandAll()

    def render_data(self,dat):
        self.m_dataVLC.DeleteAllItems()
        # for token in dat.columns:
        #     self.m_dataVLC.AppendTextColumn(token)
        for row in dat.values:
            self.m_dataVLC.AppendItem(row)
        
    def m_splitwin1OnIdle(self, event):
        # self.SetSize((WIN_DEFAULT_WIDTH,WIN_DEFAULT_HEIGHT))
        # self.m_splitwin1.SetSashPosition(0)
        self.m_splitwin1.SetSize((WIN_DEFAULT_WIDTH*0.2,WIN_DEFAULT_HEIGHT))
        self.m_tree.SetSize(self.m_splitwin1.GetSize())
        self.render_tree()
        self.m_splitwin1.Unbind(wx.EVT_IDLE)
    
    def m_splitwin2OnIdle(self, event):
        self.m_splitwin2.SetPosition((WIN_DEFAULT_WIDTH*0.2,0))
        self.m_splitwin2.SetSize((WIN_DEFAULT_WIDTH*0.6,WIN_DEFAULT_HEIGHT))
        self.m_comboBox.SetPosition((0,0))
        self.m_comboBox.SetSize((0.545*WIN_DEFAULT_WIDTH,HEIGHT_SCALE_FAC*WIN_DEFAULT_HEIGHT))
        self.m_btnSearch.SetPosition((0.545*WIN_DEFAULT_WIDTH,0))
        self.m_btnSearch.SetSize((0.05*WIN_DEFAULT_WIDTH,HEIGHT_SCALE_FAC*WIN_DEFAULT_HEIGHT))
        self.m_dataVLC.SetPosition((0,HEIGHT_SCALE_FAC*WIN_DEFAULT_HEIGHT))
        self.m_dataVLC.SetSize(self.m_splitwin2.GetSize())
        self.render_data(self.data.head(20))
        self.m_splitwin2.Unbind(wx.EVT_IDLE)
    
    def m_splitwin3OnIdle(self, event):
        self.m_splitwin3.SetPosition((WIN_DEFAULT_WIDTH*0.8,0))
        self.m_splitwin3.SetSize((WIN_DEFAULT_WIDTH*0.2,WIN_DEFAULT_HEIGHT))
        self.m_btnDetail.SetSize((0.2*WIN_DEFAULT_WIDTH,HEIGHT_SCALE_FAC*WIN_DEFAULT_HEIGHT))
        self.m_textDetail.SetPosition((0,HEIGHT_SCALE_FAC*WIN_DEFAULT_HEIGHT))
        self.m_textDetail.SetSize(self.m_splitwin3.GetSize())
        self.m_splitwin3.Unbind(wx.EVT_IDLE)

if __name__ == "__main__":
    data = pd.read_excel(DATA_PATH).set_index(
        ["分支","一级分支","二级分支"]).fillna("")
    app = wx.App()
    window = EnExcel(None,data)
    window.Show(True) 
    app.MainLoop()
