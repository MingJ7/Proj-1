class column_headers:
    userh = "User" #NOTE CHANGE THESE SHOULD THE HEADERS CHANGE
    dateh = "Date"
    changeh = "Changed"
    actionh = "Action"
    def __init__(self,row1):
        self.userCol = None #to store the column
        self.dateCol = None
        self.changeByCol = None
        self.actionCol = None
        for cell in row1:
            if cell.value is not None: #Assign the respective value of the column if the text in cell matches the header
                if self.userh in cell.value:
                    self.userCol = cell.column
                    print("{}{} is user".format(self.userCol,cell.row))
                elif self.dateh == cell.value:
                    self.dateCol = cell.column
                    print("{}{} is date".format(self.dateCol,cell.row))
                elif self.changeh in cell.value:
                    self.changeByCol = cell.column
                    print("{}{} is change".format(self.changeByCol,cell.row))
                elif self.actionh in cell.value:
                    self.actionCol = cell.column
                    print("{}{} is action".format(self.actionCol,cell.row))