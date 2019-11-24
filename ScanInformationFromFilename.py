# ScanInformationFromFilename - a script for ComicRack
# Based on Arturo's New Comics Toolbox for ComicRack (v. 0.3) Regex by Stonepaw
# Modified by Stonepaw & 600WPMPO
# New regex by the amazing Helmic
# v 0.6 (rel. 2011-November-18)


import clr, re
import System

clr.AddReference("System.Windows.Forms")

clr.AddReference("System.Drawing")

from System.IO import FileInfo

from System.Text.RegularExpressions import Regex, RegexOptions

from System.ComponentModel import BackgroundWorker

from System.Drawing import Point, Size

from System.Windows.Forms import Form, ListBox, Button, Label, TextBox, DialogResult, MessageBox, TabControl, TabPage, DockStyle, BorderStyle


#Some important constants
FOLDER = FileInfo(__file__).DirectoryName + "\\"
SCANNERSFILE = FOLDER + "scanners.txt"
BLACKLISTFILE = FOLDER + "blacklist.txt"
USERBLACKLISTFILE = FOLDER + "userblacklist.txt"
SETTINGSFILE = FOLDER + "settings.dat"
ICON = FOLDER + "ScanInformationFromFilename.ico"

#@Name Scan Information From Filename
#@Hook Books
#@Image ScanInformationFromFilename.png
#@Key scaninfofromfilename
def ScanInformationFromFilename(books):
    
    progress = ProgressDialog(books)

    progress.ShowDialog()

    progress.Dispose()


def FindScanners(worker, books):
    
    #Load the various settings. settings is a dict
    settings = LoadSettings()

    #Load the scanners
    unformatedscanners = LoadListFromFile(SCANNERSFILE)

    #Sort the scanners by length and reverse it. For example cl will come after clickwheel allowing them to be matched correctly.
    unformatedscanners.sort(key=len, reverse=True)

    #Format the scanners for use in the regex
    scanners = "|".join(unformatedscanners)
    scanners = "(?<Tags>" + scanners + ")"


    #Load the blacklist and format it
    blacklist = LoadListFromFile(BLACKLISTFILE)

    blacklist.extend(LoadUserBlackListFromFile(USERBLACKLISTFILE))

    formatedblacklist = "|".join(blacklist)


    #Add in the blacklist

    #These amazing regex are designed by the amazing Helmic.

    pattern = r"(?:(?:__(?!.*__[^_]))|[(\[])(?!(?:" + formatedblacklist + r"|[\s_\-\|/,])+[)\]])(?<Tags>(?=[^()\[\]]*[^()\[\]\W\d_])[^()\[\]]{2,})[)\]]?"

    replacePattern = r"(?:[^\w]|_|^)(?:" + formatedblacklist + r")(?:[^\w]|_|$)"

    #Create the regex

    regex = Regex(pattern, RegexOptions.IgnoreCase)
    regexScanners = Regex(scanners, RegexOptions.IgnoreCase)
    regexReplace = Regex(replacePattern, RegexOptions.IgnoreCase)

    ComicBookFields = ComicRack.App.GetComicFields()
    ComicBookFields.Remove("Scan Information")
    ComicBookFields.Add("Language", "LanguageAsText")

    for book in books:

        #.net Regex
        #Note that every possible match is found and then the last one is used.
        #This is because in some rare cases more than one thing is mistakenly matched and the scanner is almost always the last match.
        matches = regex.Matches(book.FileName)
        unknowntag = ""

        try:
            match = matches[matches.Count-1]
            
        except ValueError:
            
            #No match
            #print "Trying the Scanners.txt list"

            #Check the defined scanner names
            match = regexScanners.Match(book.FileName)

            #Still no match
            if match.Success == False:
                if settings["Unknown"] != "":
                    unknowntag = settings["Unknown"]
                else:
                    continue                


        #Check if what was grabbed is a field in the comic
        fields = []
        for field in ComicBookFields.Values:
            fields.append(unicode(getattr(book, field)).lower())

        if match.Groups["Tags"].Value.lower() in fields:
            print "Uh oh. That matched tag is in the info somewhere."
            newmatch = False
            for n in reversed(range(0, matches.Count-1)):
                if not matches[n].Groups["Tags"].Value.lower() in fields:
                    match = matches[n]
                    newmatch = True
                    break
            if newmatch == False:
                if settings["Unknown"] != "":
                    unknowntag = settings["Unknown"]
                else:
                    continue

        #Check if the match can be found in () in the series, title or altseries
        titlefields = [book.ShadowSeries, book.ShadowTitle, book.AlternateSeries]
        abort = False
        for title in titlefields:
            titleresult = re.search("\((?P<match>.*)\)", title)
            if titleresult != None and titleresult.group("match").lower() == match.Groups["Tags"].Value.lower():
                #The match is part of the title, series or altseries so skip it
                print "The match is part of the title, series or altseries"
                abort = True
                break
        if abort == True:
            if settings["Unknown"] != "":
                unknowntag = settings["Unknown"]
            else:
                continue

        #Get a list of the old ScanInformation
        oldtags = book.ScanInformation
        ListOfTagsTemp=oldtags.split(",")
        if '' in ListOfTagsTemp:
            ListOfTagsTemp.remove('')
        
        ListOfTags=[]
        if ListOfTagsTemp != []:
            for indtag in ListOfTagsTemp:
                ListOfTags.append(indtag.strip())

        #Create our new tag
        if unknowntag != "":
            newtag = settings["Prefix"] + unknowntag
        else:
            newtag = settings["Prefix"] + regexReplace.Replace(match.Groups["Tags"].Value.strip("_, "), "")
        
        if newtag not in ListOfTags:
            ListOfTags.append(newtag)

        #Sort alphabeticaly to be neat
        ListOfTags.sort()


        #Add to ScanInformation field
        book.ScanInformation = ", ".join(ListOfTags)


      

#@Key scaninfofromfilename
#@Hook ConfigScript
def ScanInformationFromFilenameOptions():
    
    settings = LoadSettings()

    scanners = LoadListFromFile(SCANNERSFILE)

    blacklist = LoadUserBlackListFromFile(USERBLACKLISTFILE)

    optionform = OptionsForm(scanners, blacklist, settings["Prefix"], settings["Unknown"])

    result = optionform.ShowDialog()

    if result == DialogResult.OK:
        settings["Prefix"] = optionform.Prefix.Text
        settings["Unknown"] = optionform.Unknown.Text
        SaveScanners(list(optionform.ScannerNames.Items))
        SaveBlackList(list(optionform.Blacklist.Items))
        SaveSettings(settings)

def LoadSettings():
    #Define some default settings
    settings = {"Prefix" : "Scanner:", "Unknown" : "Unknown"}

    #The settings file should be formated with each line as SettingName:Value. eg Prefix:Scanner:

    try:
        with open(SETTINGSFILE, 'r') as settingsfile:
            for line in settingsfile:
                match = re.match("(?P<setting>.*?):(?P<value>.*)", line)
                settings[match.group("setting")] = match.group("value")

    except Exception, ex:
        print "Something has gone wrong loading the settings file. The error was: " + str(ex)
    
    return settings

def SaveSettings(settings):
    
    with open(SETTINGSFILE, 'w') as settingsfile:
        for setting in settings:
            settingsfile.write(setting + ":" + settings[setting] + "\n")

def LoadListFromFile(filepath):
    #The file should be formated with each list item as a new line.
    #It doesn't matter what order the scanners are in the file.
    with open(filepath, 'r') as f:
        l = f.read().splitlines()

    return l


def LoadUserBlackListFromFile(filepath):
    #The file should be formated with each list item as a new line.
    #It doesn't matter what order the scanners are in the file.
    with open(filepath, 'r') as f:
        l = f.read().splitlines()
    sl = []
    for i in l:
        sl.append(re.sub(r"[\[\]\\^$.|?*+(){}]", "", i))
    return sl

def SaveScanners(scanners):
    with open(SCANNERSFILE, 'w') as scannersfile:
        for scanner in scanners:
            scannersfile.write(scanner + "\n")

def SaveBlackList(blacklist):
    with open(USERBLACKLISTFILE, 'w') as f:
        for item in blacklist:
            f.write(item + "\n")

class OptionsForm(Form):
    def __init__(self, scanners, blacklist, prefix, unknown):
        self.InitializeComponent()

        self.Prefix.Text = prefix
        self.Unknown.Text = unknown
        self.ScannerNames.Items.AddRange(System.Array[System.String](scanners))
        self.Blacklist.Items.AddRange(System.Array[System.String](blacklist))
    
    def InitializeComponent(self):
        self.ScannerNames = ListBox()
        self.Blacklist = ListBox()
        self.Add = Button()
        self.Remove = Button()
        self.Prefix = TextBox()
        self.Unknown = TextBox()
        self.lblunknown = Label()
        self.lblprefix = Label()
        self.Okay = Button()
        self.Tabs = TabControl()
        self.ScannersTab = TabPage()
        self.BlacklistTab = TabPage()
        # 
        # ScannerNames
        #
        self.ScannerNames.BorderStyle = BorderStyle.None
        self.ScannerNames.Dock = DockStyle.Fill
        self.ScannerNames.TabIndex = 0
        self.ScannerNames.Sorted = True
        # 
        # Blacklist
        # 
        self.Blacklist.Dock = DockStyle.Fill
        self.Blacklist.TabIndex = 0
        self.Blacklist.Sorted = True
        self.Blacklist.BorderStyle = BorderStyle.None
        # 
        # Add
        # 
        self.Add.Location = Point(228, 102)
        self.Add.Size = Size(75, 23)
        self.Add.Text = "Add"
        self.Add.Click += self.AddItem
        # 
        # Remove
        # 
        self.Remove.Location = Point(228, 162)
        self.Remove.Size = Size(75, 23)
        self.Remove.Text = "Remove"
        self.Remove.Click += self.RemoveItem
        # 
        # Prefix
        # 
        self.Prefix.Location = Point(76, 313)
        self.Prefix.Size = Size(136, 20)
        # 
        # lblprefix
        # 
        self.lblprefix.AutoSize = True
        self.lblprefix.Location = Point(12, 316)
        self.lblprefix.Size = Size(58, 13)
        self.lblprefix.Text = "Tag Prefix:"
        # 
        # Unknown
        # 
        self.Unknown.Location = Point(90, 337)
        self.Unknown.Size = Size(122, 20)
        # 
        # lblprefix
        # 
        self.lblunknown.AutoSize = True
        self.lblunknown.Location = Point(12, 340)
        self.lblunknown.Size = Size(58, 13)
        self.lblunknown.Text = "Unknown Tag:"
        # 
        # Okay
        # 
        self.Okay.Location = Point(228, 339)
        self.Okay.Size = Size(75, 23)
        self.Okay.Text = "Okay"
        self.Okay.DialogResult = DialogResult.OK
        #
        # ScannersTab
        #
        self.ScannersTab.Text = "Scanners"
        self.ScannersTab.UseVisualStyleBackColor = True
        self.ScannersTab.Controls.Add(self.ScannerNames)
        #
        # BlacklistTab
        #
        self.BlacklistTab.Text = "Blacklist"
        self.BlacklistTab.UseVisualStyleBackColor = True
        self.BlacklistTab.Controls.Add(self.Blacklist)
        #
        # Tabs
        #
        self.Tabs.Size = Size(210, 280)
        self.Tabs.Location = Point(12, 12)
        self.Tabs.Controls.Add(self.ScannersTab)
        self.Tabs.Controls.Add(self.BlacklistTab)
        # 
        # Form Settings
        # 
        self.Size = System.Drawing.Size(315, 400)
        self.Controls.Add(self.Tabs)
        self.Controls.Add(self.Add)
        self.Controls.Add(self.Remove)
        self.Controls.Add(self.lblprefix)
        self.Controls.Add(self.Prefix)
        self.Controls.Add(self.lblunknown)
        self.Controls.Add(self.Unknown)
        self.Controls.Add(self.Okay)
        self.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        self.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        self.Text = "Scan Information From Filename Options"
        self.MinimizeBox = False
        self.MaximizeBox = False
        self.AcceptButton = self.Okay
        self.Icon = System.Drawing.Icon(ICON)


    def RemoveItem(self, sender, e):
        if self.Tabs.SelectedTab == self.ScannersTab:
            self.ScannerNames.Items.Remove(self.ScannerNames.SelectedItem)
        else:
            self.Blacklist.Items.Remove(self.Blacklist.SelectedItem)

    def AddItem(self, sender, e):
        input = InputBox()
        input.Owner = self
        if input.ShowDialog() == DialogResult.OK:
            if self.Tabs.SelectedTab == self.ScannersTab:
                self.ScannerNames.Items.Add(input.FindName())
                self.ScannerNames.SelectedItem = input.FindName()
            else:
                self.Blacklist.Items.Add(input.FindName())
                self.Blacklist.SelectedItem = input.FindName()                

class InputBox(Form):
	def __init__(self):
		self.TextBox = TextBox()
		self.TextBox.Size = Size(250, 20)
		self.TextBox.Location = Point(15, 12)
		self.TextBox.TabIndex = 1
		
		self.OK = Button()
		self.OK.Text = "OK"
		self.OK.Size = Size(75, 23)
		self.OK.Location = Point(109, 38)
		self.OK.DialogResult = DialogResult.OK
		self.OK.Click += self.CheckTextBox
		
		self.Cancel = Button()
		self.Cancel.Size = Size(75, 23)
		self.Cancel.Text = "Cancel"
		self.Cancel.Location = Point(190, 38)
		self.Cancel.DialogResult = DialogResult.Cancel
		
		self.Size = Size(300, 100)
		self.Text = "Please enter a scanner name"
		self.Controls.Add(self.OK)
		self.Controls.Add(self.Cancel)
		self.Controls.Add(self.TextBox)
		self.AcceptButton = self.OK
		self.CancelButton = self.Cancel
		self.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		self.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
		self.Icon = System.Drawing.Icon(ICON)
		self.ActiveControl = self.TextBox
		
	def FindName(self):
		if self.DialogResult == DialogResult.OK:
			return self.TextBox.Text.strip()
		else:
			return None
		
	def CheckTextBox(self, sender, e):
		if not self.TextBox.Text.strip():
			MessageBox.Show("Please enter a name into the textbox")
			self.DialogResult = DialogResult.None
		
		if self.TextBox.Text.strip() in self.Owner.ScannerNames.Items:
			MessageBox.Show("The entered name is already in entered. Please enter another")
			self.DialogResult = DialogResult.None

class ProgressDialog(Form):
    
    def __init__(self, books):
        self.InitializeComponent()
        self.worker.RunWorkerAsync(books)
        self.done = False

    def InitializeComponent(self):
        self.progressBar = System.Windows.Forms.ProgressBar()
        # 
        # progressBar
        # 
        self.progressBar.Location = Point(0, 0)
        self.progressBar.Size = Size(284, 23)
        self.progressBar.Maximum = 100
        self.progressBar.Style = System.Windows.Forms.ProgressBarStyle.Marquee

        self.worker = BackgroundWorker()
        self.worker.DoWork += self.WorkerDoWork
        self.worker.RunWorkerCompleted += self.WorkerCompleted

        self.ClientSize = Size(284, 23)
        self.Controls.Add(self.progressBar)
        self.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.Text = "Searching Filenames"
        self.Icon = System.Drawing.Icon(ICON)
        self.FormClosing += self.CheckClosing
        self.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen

    #Stop the user from closing the progress dialog
    def CheckClosing(self, sender, e):
        if e.CloseReason == System.Windows.Forms.CloseReason.UserClosing and self.done == False:
            e.Cancel = True

    def SetTitle(self, text):
        self.Text = text

    def WorkerDoWork(self, sender, e):
        FindScanners(sender, e.Argument)
        e.Result = "Done"

    def WorkerCompleted(self, sender, e):
        self.done = True
        self.Close()
