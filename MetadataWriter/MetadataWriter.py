import os, sys
import openpyxl
import exiftool
from tkinter import filedialog
from pprint import pprint

class MetadataWriter():
    
    def __init__(self, *args, **kwargs):
        self.fn = filedialog.askopenfilename(title='Open', 
                                             defaultextension='.xlsx', filetypes=[('Excel Files', '*.xlsx')])
        head, tail = os.path.split(self.fn)
        self.directory = head
        self.f = openpyxl.load_workbook(self.fn)
        self.et = exiftool.ExifTool('C:\\Users\\evharley\\Downloads\\exiftool-11.32\\exiftool.exe')
        self.ws = self.f.active
        
        
    def parse_workbook(self):
        self.et.start()
        keys = {self.ws[1][col].value: col for col in range(self.ws.max_column)}
        list_of_fields_used = []
        for row in range(2, self.ws.max_row):
            folder = self.ws[row][keys['Directory']].value
            filename = self.ws[row][keys['file_name']].value
            filepath = '{0}/{1}{2}'.format(self.directory,folder, filename)
            metadata = self.et.get_metadata(filepath)
            xmp_data = {key: metadata[key] for key in metadata.keys() if key.startswith('XMP:')
                        and not key.startswith('XMP:F')}
            metadata_to_remove =  ['XMP:About', 'XMP:Cache', 'XMP:Checkout', 'XMP:Colorprofile',
                                   'XMP:Directory_id', 'XMP:Discussion_count', 'XMP:DocumentID',
                                   'XMP:Duration', 'XMP:Manager', 'XMP:ManagerVariant', 'XMP:Mb_id',
                                   'XMP:Needs_xmp_auto', 'XMP:Orig_x', 'XMP:Orig_y', 'XMP:Page_count',
                                   'XMP:Rotate', 'XMP:Thumbnails_lock', 'XMP:Thumbnails_x',
                                   'XMP:Thumbnails_y', 'XMP:Usermodified', 'XMP:Version_of', 
                                   'XMP:Video_status', 'XMP:View_sched', 'XMP:Viewex_lock',
                                   'XMP:Viewex_y', 'XMP:XMPToolkit', 'XMP:Xmp_volatile', 'XMP:Zoom']
            for key in metadata_to_remove:
                xmp_data.pop(key, None)

            keyword_keys = ['XMP:Caption', 'XMP:Topicresponsibility', 'XMP:NeedsData', 'XMP:TopicConservation',
                                'XMP:CallNumber']
            author_keys = ['XMP:Author']
            for key in xmp_data.keys():
                if key in keyword_keys:
                    keywords += xmp_data[key]


        pprint(list_of_fields_used)   
        return 0
    
   