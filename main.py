import os
from win32com.client import Dispatch, constants, gencache, DispatchEx
from PyPDF2 import PdfFileWriter, PdfFileReader, PdfFileMerger
import tkinter as tk
from tkinter import filedialog

def getLocalFile():
    #获取选择文件的路径
    root = tk.Tk()
    root.withdraw()
    filePath = filedialog.askopenfilename()
    print('file path: ', filePath)
    return filePath

def getLocalFolder():
    #获取选择文件的路径
    root = tk.Tk()
    root.withdraw()
    filePath = filedialog.askdirectory()
    print('folder path: ', filePath)
    return filePath

class PDFConverter:
    def __init__(self, pathname, export='.'):
        self._handle_postfix = ['doc', 'docx', 'ppt', 'pptx', 'xls', 'xlsx']
        self._filename_list = list()
        name = os.path.basename(pathname).split('.')[0]
        self._export_folder = os.path.join(os.path.abspath('.'), name)
        if not os.path.exists(self._export_folder):
            os.mkdir(self._export_folder)
        self._enumerate_filename(pathname)

    def _enumerate_filename(self, pathname):
        '''
        读取所有文件名
        '''
        full_pathname = os.path.abspath(pathname)
        if os.path.isfile(full_pathname):
            if self._is_legal_postfix(full_pathname):
                self._filename_list.append(full_pathname)
            else:
                raise TypeError('文件 {} 后缀名不合法！仅支持如下文件类型：{}。'.format(pathname, '、'.join(self._handle_postfix)))
        elif os.path.isdir(full_pathname):
            for relpath, _, files in os.walk(full_pathname):
                for name in files:
                    filename = os.path.join(full_pathname, relpath, name)
                    if self._is_legal_postfix(filename):
                        self._filename_list.append(os.path.join(filename))
        else:
            raise TypeError('文件/文件夹 {} 不存在或不合法！'.format(pathname))

    def _is_legal_postfix(self, filename):
        return filename.split('.')[-1].lower() in self._handle_postfix and not os.path.basename(filename).startswith(
            '~')

    def run_conver(self):
        '''
        进行批量处理，根据后缀名调用函数执行转换
        '''
        print('需要转换的文件数：', len(self._filename_list))
        for filename in self._filename_list:
            postfix = filename.split('.')[-1].lower()
            funcCall = getattr(self, postfix)
            print('原文件：', filename)
            funcCall(filename)
        print('转换完成！')

    def doc(self, filename):
        '''
        doc 和 docx 文件转换
        '''
        name = os.path.basename(filename).split('.')[0] + '.pdf'
        exportfile = os.path.join(self._export_folder, name)
        print('保存 PDF 文件：', exportfile)
        gencache.EnsureModule('{00020905-0000-0000-C000-000000000046}', 0, 8, 4)
        w = Dispatch("Word.Application")
        doc = w.Documents.Open(filename)
        doc.ExportAsFixedFormat(exportfile, constants.wdExportFormatPDF,
                                Item=constants.wdExportDocumentWithMarkup,
                                CreateBookmarks=constants.wdExportCreateHeadingBookmarks)

        w.Quit(constants.wdDoNotSaveChanges)

    def docx(self, filename):
        self.doc(filename)

    def xls(self, filename):
        '''
        xls 和 xlsx 文件转换
        '''
        name = os.path.basename(filename).split('.')[0] + '.pdf'
        exportfile = os.path.join(self._export_folder, name)
        xlApp = DispatchEx("Excel.Application")
        xlApp.Visible = False
        xlApp.DisplayAlerts = 0
        books = xlApp.Workbooks.Open(filename, False)
        books.ExportAsFixedFormat(0, exportfile)
        books.Close(False)
        print('folder:', self._export_folder)
        print('保存 PDF 文件：', exportfile)
        xlApp.Quit()


        inputpdf = PdfFileReader(open(exportfile, "rb"))
        for i in range(inputpdf.numPages):
            output = PdfFileWriter()
            output.addPage(inputpdf.getPage(i))
            with open(os.path.join(self._export_folder, "%s.pdf" % (i+1)), "wb") as outputStream:
                print('正在生成第%s页'%(i+1))
                output.write(outputStream)

    def xlsx(self, filename):
        self.xls(filename)

    def ppt(self, filename):
        '''
        ppt 和 pptx 文件转换
        '''
        name = os.path.basename(filename).split('.')[0] + '.pdf'
        exportfile = os.path.join(self._export_folder, name)
        gencache.EnsureModule('{00020905-0000-0000-C000-000000000046}', 0, 8, 4)
        p = Dispatch("PowerPoint.Application")
        ppt = p.Presentations.Open(filename, False, False, False)
        ppt.ExportAsFixedFormat(exportfile, 2, PrintRange=None)
        print('保存 PDF 文件：', exportfile)
        p.Quit()

    def pptx(self, filename):
        self.ppt(filename)

    def getTargetFolder(self):
        return self._export_folder


class Merge2Folder:
    def __init__(self, folder1, folder2):
        self.folder1 = folder1 # abs path
        self.folder2 = folder2

    def merge2file(self, t1, t2, location):
        if not os.path.exists(location):
            os.mkdir(location)

        f1 = PdfFileReader(open(t1, "rb"))
        f2 = PdfFileReader(open(t2, "rb"))
        mergedObject = PdfFileMerger()
        mergedObject.append(f1)
        mergedObject.append(f2)
        name = os.path.basename(t1).split('.')[0]
        res = os.path.join(location, 'merged_'+name+'.pdf')
        mergedObject.write(open(res, 'wb'))
        print('{}与{}合并完成'.format(t1, t2))

    def merge2Folder(self):
        # list of files in folder1
        dirs1 = os.listdir(self.folder1) # only name

        # files in folder2
        dirs2 = os.listdir(self.folder2)

        # check the format
        paired_files = self.checkMatchFile(dirs1, dirs2)
        path = os.getcwd()
        loaction = os.path.join(path, '合并后的文件')
        for files in paired_files:
            self.merge2file(files[0], files[1], loaction)



    def checkMatchFile(self, d1, d2):
        # return a list of pairs of abs path of two files that are going to merge

        # make sure all files in d2 are pdf
        d2.sort()
        for f2 in d2:
            if f2.split('.')[1] != 'pdf':
                raise TypeError('文件 {} 不是pdf格式'.format(f2))
        # make sure all files in d2 has corresponding file in d1 and return a list of pair
        res = []
        for f2 in d2:
            if f2 not in d1:
                raise TypeError('文件 {} 在 {} 中找不到匹配文件'.format(f2, self.folder1))
            else:
                res.append((os.path.join(self.folder1, f2), os.path.join(self.folder2, f2)))
        return res







if __name__ == "__main__":
    # 支持文件夹批量导入
    print('选择要转换的excel')
    pathname = getLocalFile()
    print(pathname)
    pdfConverter = PDFConverter(pathname)
    pdfConverter.run_conver()

    print('选择要合并的文件夹')
    target = pdfConverter.getTargetFolder()
    otherFolder = getLocalFolder()
    pdfMerger = Merge2Folder(target, otherFolder)
    pdfMerger.merge2Folder()
    print('合并完成！ 不用谢我，谢谢朱越 ^_^')

# 把excel转成pdf，pdf总文件与每页文件保存在自动生成的原excel名称的文件夹里（封面文件夹），每个pdf文件（封面）的名称为他在excel中对应的页数，文件夹与该exe程序处于同一路径，建议把该程序保存至桌面这样文件夹也在桌面了
# 选择里一个本地文件夹，该文件夹内的所有文件名必须与他的封面名称一致，该程序会在封面文件夹里找到每个本地文件的同名文件并作为他的封面生成新文件保存在’合并后的文件‘文件夹内