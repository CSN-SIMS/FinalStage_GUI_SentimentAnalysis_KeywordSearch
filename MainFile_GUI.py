# Rosen
# GUI with functionality connecting the two parts of the program - keyword-search Categorization and Sentiment Analysis

from tkinter import Tk, Button, Frame, PhotoImage, Message, Canvas, Label, Listbox, Scrollbar, \
    RIGHT, X, Y, END, BOTTOM, HORIZONTAL, VERTICAL, Entry, Checkbutton, OptionMenu, Toplevel, filedialog, StringVar, IntVar
from KeywordSearch import *
from Functions import *
from Preperation import *
from PrepFunctions import *
from Analysis import *

FontStyle = "Helvetica"

# Page frame for switching between the three different pages of the application
class Page(Frame):
    def __init__(self, *args, **kwargs):
        Frame.__init__(self, *args, **kwargs)
    def show(self):
        self.lift()
# Background of the application
def backgroundSet(self):
    self.canvas = Canvas(self)
    self.background_label = Label(self, bg='#3C1E5F')
    self.background_label.place(relwidth=1, relheight=1)
    self.canvas.pack()

# Popup message with result from page 1
# Gets the selected folder by the user and uses keywordSearch in txt files, then presents categories and file names
# Sentiment analysis of the output sub-directories and saving the result in excel
def popupWindowInputFiles(selectLanguageVar, checkVar, entryContentExcelFilename):
    print("Language value " + str(selectLanguageVar))
    print("Check value " + str(checkVar))
    print("Entry excel name " + str(entryContentExcelFilename))
    errorText = ""
    amountOfFiles = 0
    mapCategorizedFiles = {}
    messageSavingExcelFile = ""
    noExcel = True
    judgementList = []
    try:
        # Ask for directory with input files
        folderSelected = filedialog.askdirectory()
        categorizer = Categorizer(folderSelected)
        # Amount of analyzed Emails
        amountOfFiles = categorizer.amountOfFiles(folderSelected)
        # Map of categories with their emails
        mapCategorizedFiles = categorizer.categorizeFilesFromDirectoryInMapAndSubDirectory()
        # Preparation of Sentiment analysis
        if (selectLanguageVar == 'Swedish'):
            prepareAnalysis(True)
        else:
            prepareAnalysis(False)
        # Start of Sentiment analysis
        if(checkVar==1 and entryContentExcelFilename != ''):
            noExcel = False
            judgementList = startAnalysis(checkVar, entryContentExcelFilename, noExcel)
            messageSavingExcelFile = "The result of Sentiment analysis and categorization " + "\n" + " is saved as " + str(entryContentExcelFilename) + ".xlsx"
        elif (checkVar == 1 and entryContentExcelFilename == ''):
            noExcel = False
            judgementList = startAnalysis(checkVar, entryContentExcelFilename, noExcel)
            messageSavingExcelFile = "The result of Sentiment analysis and categorization " + "\n" + " is saved as newTestfile.xlsx"
        elif (checkVar == 0 and entryContentExcelFilename == ''):
            judgementList = startAnalysis(checkVar, entryContentExcelFilename, noExcel)
            messageSavingExcelFile = "Please state a name of the new excel file in the entry field"
        elif (checkVar == 0 and entryContentExcelFilename != ''):
            judgementList = startAnalysis(checkVar, entryContentExcelFilename, noExcel)
            messageSavingExcelFile = "Please check the box for new excel file in the entry field"
    except UnicodeDecodeError:
        errorText = "Selected folder does not contain txt files."
    except FileNotFoundError:
        errorText = "The system cannot find the path specified."
    popupWindowInputFiles = Toplevel()
    popupWindowInputFiles.wm_title("Result")
    popupWindowInputFiles.wm_geometry("800x450")
    # container with results from Categorization and Sentiment analysis of the given input folder
    results = Listbox(popupWindowInputFiles, font=("Courier", 12), bg='white', fg='#3C1E5F', justify='left', bd=3)
    results.grid(column=1, row=1, padx=10, ipady=10)
    results.place(relwidth=1, relheight=0.85)
    scrollbar_vertical = Scrollbar(results, orient=VERTICAL)
    scrollbar_vertical.pack(side=RIGHT, fill=Y)
    scrollbar_vertical.configure(command=results.yview)
    scrollbar_horizontal = Scrollbar(results, orient=HORIZONTAL)
    scrollbar_horizontal.pack(side=BOTTOM, fill=X)
    scrollbar_horizontal.configure(command=results.xview)
    results.configure(yscrollcommand=scrollbar_vertical.set)
    results.configure(xscrollcommand=scrollbar_horizontal.set)

    # Shows if there is an error occurred when opening the input folder
    if(errorText != ""):
        results.insert(END, "Error occured: " + errorText)
        results.insert(END, "Try again selecting an input folder with text files")
        results.insert(END, "\n")
    else:
        # Shows the amount of analyzed Emails
        results.insert(END, "Amount of analysed Emails: " + str(amountOfFiles))
        results.insert(END, "\n")

        # Shows a map of categories with their emails
        results.insert(END, "List of categories with their emails: ")
        results.insert(END, "Category".ljust(20, ' ') + "File name")
        for key, val in mapCategorizedFiles.items():
            results.insert(END, str(key).ljust(20, ' ') + str(val))

        # Shows the message about saving the result as excel file
        results.insert(END, "\n")
        results.insert(END, messageSavingExcelFile)
        results.insert(END, "\n")

        # Shows the result from judgementList with Filename, Category, Judgement from Sentiment analysis in %, Confidence
        results.insert(END, "Below is the result: \n")
        results.insert(END, "Filename".ljust(30, ' ') + "Category".ljust(20, ' ')
                       + "Judgement %".ljust(15, ' ') + "Confidence")
        for fn, judgement in judgementList:
            results.insert(END, str(fn[0]).ljust(30, ' ') + str(fn[1]).ljust(20, ' ')
                           + str(judgement[0]).ljust(15, ' ') + str(judgement[1]))

    # Button to close the popup window
    buttonPopup = Button(popupWindowInputFiles, text="Okay", bd=4, command=popupWindowInputFiles.destroy)
    buttonPopup.place(relx=0.4, rely=0.85, relwidth=0.2, relheight=0.15)

# Popup message with result from Sentiment analysis of direct input by the user
def popupWindowDirectInput(selectLanguageVar, entryString):
    print(selectLanguageVar)
    print(entryString)
    # Preparation of Sentiment analysis of an entry by the user
    if (selectLanguageVar == 'Swedish'):
        translatedMessageList = translateEntryMessageToEnglish(entryString)
        savePickle(translatedMessageList, "picklefiles_eng/translatedmessages.pickle")
    else:
        savePickle(entryString, "picklefiles_eng/messages.pickle")
    # Start of Sentiment analysis of an entry by the user
    print(sentiment(entryString, voted_classifier, word_features))
    resultSentimentAndConfidence = sentiment(entryString, voted_classifier, word_features)
    # Popup message with result from Sentiment analysis of direct input by the user (displays positive, negative or neutral sentiment)
    popupWindow = Toplevel()
    popupWindow.wm_title("Result")
    popupWindow.wm_geometry("500x400")
    resultLabel = Label(popupWindow, font=(FontStyle, 16))
    resultLabel.place(relx=0.32, rely=0.1, relwidth=0.35, relheight=0.15)
    resultSentiment = StringVar()
    if(str(resultSentimentAndConfidence[0]) == "pos"):
        resultSentiment = "Positive"
        resultLabel.configure(text=resultSentiment, fg='#00FF00')
    elif (str(resultSentimentAndConfidence[0]) == "neg"):
        resultSentiment = "Negative"
        resultLabel.configure(text=resultSentiment, fg='#FF0000')
    else:
        resultSentiment = "Neutral"
        resultLabel.configure(text=resultSentiment, fg='#808080')
    resultMessage = "Your entry \"" + entryString + "\" has " + resultSentiment + " Sentiment and Confidence " + str(resultSentimentAndConfidence[1])
    infoMessage = Message(popupWindow, text=str(resultMessage), font=(FontStyle, 14), justify='center', aspect=150)
    infoMessage.place(relx=0.05, rely=0.2, relwidth=0.9, relheight=0.5)
    # Button to close the popup window
    buttonPopup = Button(popupWindow, text="Okay", bd=4, command=popupWindow.destroy)
    buttonPopup.place(relx=0.4, rely=0.75, relwidth=0.2, relheight=0.15)

# First page with main information about the application
class Page1(Page):
   def __init__(self, *args, **kwargs):
       Page.__init__(self, *args, **kwargs)
       backgroundSet(self)
       # Lower frame with scrollbars for displaying of information about the program
       lower_frame = Frame(self, bg='#FFD164', bd=5)
       lower_frame.place(relx=0.5, rely=0.15, relwidth=0.9, relheight=0.7, anchor='n')
       lower_frame.grid_rowconfigure(0, weight=1)
       lower_frame.grid_columnconfigure(0, weight=1)
       infoMessagePage1 = Listbox(lower_frame, font=(FontStyle, 14), bg='white', fg='#3C1E5F', justify='center', bd=3)
       infoMessagePage1.grid(column=1, row=1, padx=10, ipady=10)
       infoMessagePage1.place(relwidth=1, relheight=1)
       scrollbar_vertical = Scrollbar(lower_frame, orient=VERTICAL)
       scrollbar_vertical.pack(side=RIGHT, fill=Y)
       scrollbar_vertical.configure(command=infoMessagePage1.yview)
       scrollbar_horizontal = Scrollbar(lower_frame, orient=HORIZONTAL)
       scrollbar_horizontal.pack(side=BOTTOM, fill=X)
       scrollbar_horizontal.configure(command=infoMessagePage1.xview)
       infoMessagePage1.configure(yscrollcommand=scrollbar_vertical.set)
       infoMessagePage1.configure(xscrollcommand=scrollbar_horizontal.set)
       infoMessagePage1.insert(END, "\n")
       infoMessagePage1.insert(END, "\n")
       infoMessagePage1.insert(END, "This program has the following abilities: \n")
       infoMessagePage1.insert(END, "\n")
       infoMessagePage1.insert(END, "- \"Options\" menu does Categorization and Sentiment analysis on text files in \n")
       infoMessagePage1.insert(END, "\t Swedish or English from a given input folder, \n")
       infoMessagePage1.insert(END, "\t presents the results and saves them in excel file \n\n")
       infoMessagePage1.insert(END, "- \"Direct Input\" does Sentiment analysis on direct input in Swedish or English")
       infoMessagePage1.insert(END, "\n")
       infoMessagePage1.insert(END, "\n")
       infoMessagePage1.insert(END, "\n")
       infoMessagePage1.insert(END, "\n Copyright © All rights reserved")

# Page with options for choosing input files, saving as new excel file or changing between Swedish and English
class Page2(Page):
   def __init__(self, *args, **kwargs):
       Page.__init__(self, *args, **kwargs)
       backgroundSet(self)
       # Lower frame with scrollbars for displaying of categories and file names
       lower_frame = Frame(self, bg='#FFD164', bd=5)
       lower_frame.place(relx=0.5, rely=0.15, relwidth=0.9, relheight=0.7, anchor='n')
       lower_frame.grid_rowconfigure(0, weight=1)
       lower_frame.grid_columnconfigure(0, weight=1)
       optionCanvas = Canvas(lower_frame, bg='white', bd=3)
       optionCanvas.place(relwidth=1, relheight=1)

       # select language (English or Swedish)
       selectLanguageVar = StringVar()
       # Dictionary with options
       choices = {'English', 'Swedish'}
       selectLanguageVar.set('Swedish')  # set the default option
       # Popup menu with languages
       popupMenu = OptionMenu(optionCanvas, selectLanguageVar, *choices)
       popupLabel = Label(optionCanvas, text="Choose a language", font=(FontStyle, 12), bg='white')
       popupLabel.place(relx=0.1, rely=0.01, relwidth=0.3, relheight=0.2)
       popupMenu.configure(bd=3, bg='#EE7C7D')
       popupMenu.place(relx=0.5, rely=0.01, relwidth=0.3, relheight=0.15)

       # save result in excel file
       checkVar = IntVar()
       excelFileCheckbutton = Checkbutton(optionCanvas, text="Save as excel", variable=checkVar, onvalue=1,
                                          offvalue=0, bg='white', font=(FontStyle, 12), height=5, width=20)
       excelFileCheckbutton.place(relx=0.1, rely=0.35, relwidth=0.3, relheight=0.15)
       entryLabel = Label(optionCanvas, text="Enter name of the excel", bg='white', font=(FontStyle, 12))
       entryLabel.place(relx=0.4, rely=0.38, relwidth=0.25, relheight=0.1)
       entryContentExcelFilename = Entry(lower_frame, font=(FontStyle, 12,), justify='left', bd=3)
       timestamp = datetime.datetime.now().strftime('%Y_%m_%d_%H%M%S')
       timestamp = str(timestamp)
       entryContentExcelFilename.insert(END, 'OutputFile'+timestamp)
       entryContentExcelFilename.place(relx=0.7, rely=0.38, relwidth=0.25, relheight=0.1)

       # open folder with input files
       openFolder = Label(optionCanvas, text="Open a folder with input files", justify='left',
                          bg='white', font=(FontStyle, 12))
       openFolder.place(relx=0.05, rely=0.7, relwidth=0.4, relheight=0.2)
       buttonOpenFolder = Button(optionCanvas, text="Browse", font=(FontStyle, 12), bg='#EE7C7D', highlightcolor='#d65859', activebackground='#f2d9e6',
                                 command=lambda: popupWindowInputFiles(selectLanguageVar.get(), checkVar.get(), entryContentExcelFilename.get()))
       buttonOpenFolder.place(relx=0.5, rely=0.7, relwidth=0.3, relheight=0.15)

# Page with entry field for direct input and changing between Swedish and English
class Page3(Page):
   def __init__(self, *args, **kwargs):
       Page.__init__(self, *args, **kwargs)
       backgroundSet(self)
       lower_frame = Frame(self, bg='#FFD164', bd=5)
       lower_frame.place(relx=0.5, rely=0.15, relwidth=0.9, relheight=0.7, anchor='n')
       lower_frame.grid_rowconfigure(0, weight=1)
       lower_frame.grid_columnconfigure(0, weight=1)

       # select language (English or Swedish)
       selectLanguageVar = StringVar()
       # Dictionary with options
       choices = {'English', 'Swedish'}
       selectLanguageVar.set('Swedish')  # set the default option
       # Popup menu with languages
       popupMenu = OptionMenu(lower_frame, selectLanguageVar, *choices)
       popupLabel = Label(lower_frame, text="Choose a language", font=(FontStyle, 12), bg='#FFD164')
       popupLabel.place(relx=0.1, rely=0.01, relwidth=0.3, relheight=0.2)
       popupMenu.configure(bd=3, bg='#EE7C7D')
       popupMenu.place(relx=0.5, rely=0.01, relwidth=0.3, relheight=0.15)

       # Entry text box
       entryLabel = Label(lower_frame, text="Enter text", bg='#FFD164', font=(FontStyle, 12))
       entryLabel.place(relx=0.37, rely=0.2, relwidth=0.25, relheight=0.1)
       entryContent = Entry(lower_frame, font=(FontStyle, 12,), justify='center')
       entryContent.place(relx=0, rely=0.3, relwidth=1, relheight=0.55)
       buttonAnalyzeInput = Button(lower_frame, text="Analyze", font=(FontStyle, 12), bg='#b3b3b3',
                       activebackground='#f2d9e6', command=lambda: popupWindowDirectInput(selectLanguageVar.get(), entryContent.get()))
       buttonAnalyzeInput.place(relx=0.4, rely=0.89, relwidth=0.2, relheight=0.1)

# Main view of the application with logo, info-box, button-bar and container for switching between the three pages
class MainView(Frame):
    def __init__(self, *args, **kwargs):
        Frame.__init__(self, *args, **kwargs)
        p1 = Page1(self)
        p2 = Page2(self)
        p3 = Page3(self)

        menu_frame = Frame(self, bg='#FFD164', bd=5)
        menu_frame.place(relx=0, rely=0, relwidth=1, relheight=0.2)
        #logo
        self.logo = Canvas(menu_frame, bd=1)
        self.logo.place(relx=0, rely=0, relwidth=1, relheight=0.8)
        self.img = PhotoImage(file="logo.png")
        self.img = self.img.subsample(6)
        self.logo.create_image(0, 0, anchor='nw', image=self.img)
        var = "Sentiment analysis and Categorization"
        infoMessage = Message(menu_frame, text=var, justify='center', width=350,
                                   font=(FontStyle, 16))
        infoMessage.place(relx=0.4, rely=0.1, relwidth=0.4, relheight=0.5)
        button_frame = Frame(self, bg='#FFD164', bd=5)
        button_frame.place(relx=0, rely=0.135, relwidth=1, relheight=0.3)
        button1 = Button(button_frame, text="Information", font=(FontStyle, 14), bg='#EE7C7D',
                         activebackground='#f2d9e6',
                         command=p1.lift)
        button1.place(relx=0.1, rely=0.25, relwidth=0.2, relheight=0.2)
        button2 = Button(button_frame, text="Options", font=(FontStyle, 14), bg='#EE7C7D',
                         activebackground='#f2d9e6',
                         command=p2.lift)
        button2.place(relx=0.4, rely=0.25, relwidth=0.2, relheight=0.2)
        button3 = Button(button_frame, text="Direct input", font=(FontStyle, 14), bg='#EE7C7D',
                         activebackground='#f2d9e6',
                         command=p3.lift)
        button3.place(relx=0.7, rely=0.25, relwidth=0.2, relheight=0.2)

        container = Frame(self)
        container.place(relx=0, rely=0.3, relwidth=1, relheight=0.7)

        p1.place(in_=container, x=0, y=0, relwidth=1, relheight=1)
        p2.place(in_=container, x=0, y=0, relwidth=1, relheight=1)
        p3.place(in_=container, x=0, y=0, relwidth=1, relheight=1)


        p1.show()

# Start of the application
if __name__ == "__main__":
    root = Tk()
    main = MainView(root)
    main.pack(side="top", fill="both", expand=True)
    root.title("Sentiment Classification/Categorization Dialog Widget")
    root.minsize(850, 650)
    root.mainloop()