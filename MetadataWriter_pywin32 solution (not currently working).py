    def parse_workbook(self):
        keys = {self.ws[1][i].value: i for i in range(self.ws.max_column)}
        metadata_cols = {'Caption': storagecon.PIDSI_KEYWORDS,
                         'creator': storagecon.PIDSI_AUTHOR,
                         'rights': storagecon.PIDSI_COMMENTS,
                         'subject': storagecon.PIDSI_KEYWORDS,
                         'Author': storagecon.PIDSI_AUTHOR,
                         'Comment': storagecon.PIDSI_COMMENTS,
                         'Status': storagecon.PIDSI_KEYWORDS,
                         'NeedsData': storagecon.PIDSI_KEYWORDS,
                         'Conservation Specific': storagecon.PIDSI_KEYWORDS,
                         'Catalogue Number': storagecon.PIDSI_KEYWORDS,
                         'Conservation Catalogue Sets': storagecon.PIDSI_KEYWORDS}
        for row in range(2, self.ws.max_row):
            folder = self.ws[row][keys['Directory']].value
            filename = self.ws[row][keys['file_name']].value
            filepath = '{0}/{1}{2}'.format(self.directory,folder, filename)
            metadata = {}
            pss = pythoncom.StgOpenStorageEx(filepath, STORAGE_READ, storagecon.STGFMT_FILE, 0 , pythoncom.IID_IPropertySetStorage)
            pssum = pss.Open(pythoncom.FMTID_SummaryInformation, STORAGE_READ)
            for key in metadata_cols.keys():
                data = []
                if key not in ('creator', 'Author'):
                    data.append(self.ws[row][keys[key]].value)
                elif metadata_cols[key] == 'PIDSI_AUTHOR':
                    data.append('{0}: {1}'.format(key.capitalize, self.ws[row][keys[key]].value))          
                pssum.WriteMultiple(metadata_cols[key], data)

            for name, property in self.property_sets(filepath=filepath):
                metadata[name] = property
            if row == 2:
                pprint(metadata)
        return 0
    
    def property_dict (self, property_set_storage, fmtid):
        properties = {}
        try:
            property_storage = property_set_storage.Open (fmtid, STORAGE_READ)
        except pythoncom.com_error as error:
            if error.strerror == 'STG_E_FILENOTFOUND':
                return {}
            else:
                raise
      
        for name, property_id, vartype in property_storage:
            if name is None:
                name = PROPERTIES.get (fmtid, {}).get (property_id, None)
            if name is None:
                name = hex (property_id)
            try:
                for value in property_storage.ReadMultiple ([property_id]):
                    properties[name] = value
        #
        # There are certain values we can't read; they
        # raise type errors from within the pythoncom
        # implementation, thumbnail
        #
            except TypeError:
                properties[name] = None
        return properties
  
    def property_sets(self, filepath):
        pidl, flags = shell.SHILCreateFromPath (os.path.abspath (filepath), 0)
        property_set_storage = shell.SHGetDesktopFolder ().BindToStorage (pidl, None, pythoncom.IID_IPropertySetStorage)
        for fmtid, clsid, flags, ctime, mtime, atime in property_set_storage:
            yield FORMATS.get (fmtid, str(fmtid)), self.property_dict (property_set_storage, fmtid)
        if fmtid == pythoncom.FMTID_DocSummaryInformation:
            fmtid = pythoncom.FMTID_UserDefinedProperties
            user_defined_properties = self.property_dict (property_set_storage, fmtid)
            if user_defined_properties:
                yield FORMATS.get (fmtid, str(fmtid)), user_defined_properties