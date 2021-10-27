import os
import Tkinter, Tkconstants, tkFileDialog
#import PIL.Image
#import PIL.ImageTk
from notebook import *
import time
import threading
import tkFont
import xlsxwriter


def fileselection(i):

    """Returns an opened file in read mode."""
    
    file_opt = options = {}
    options['defaultextension'] = '.txt'
    options['filetypes'] = [('all files', '.*'), ('text files', '.txt')]
    options['initialdir'] = 'C:\\'
    options['initialfile'] = 'myfile.txt'
    options['parent'] = form
    options['title'] = 'This is a title'
    
    file_name=tkFileDialog.askopenfilename(**file_opt)
    
    global config
    #global currentfile
    config.write("File:"+str(i)+" :\n"+file_name)
    config.write("\n")
    
    if len(file_name)>100:
	    file_name_short="..."+file_name[(len(file_name)-100):len(file_name)]
    else:
	    file_name_short=file_name
    
    current_button=Tkinter.Button(frames[i], text="Selected file:"+file_name_short,state=DISABLED)
    current_button.grid(row=2, column=2, sticky='EW', padx=8, pady=2)
    
    
    
def f06selection(i):

    """Returns an opened file in read mode."""
    
    file_opt = options = {}
    options['defaultextension'] = '.txt'
    options['filetypes'] = [('all files', '.*'), ('text files', '.txt')]
    options['initialdir'] = 'C:\\'
    options['initialfile'] = 'myfile.txt'
    options['parent'] = form
    options['title'] = 'This is a title'
    
    file_name=tkFileDialog.askopenfilename(**file_opt)
    
    global config
    #global currentfile
    config.write("File:"+str(i)+" :\n"+file_name)
    config.write("\n")
    
    if len(file_name)>100:
	    file_name_short="..."+file_name[(len(file_name)-100):len(file_name)]
    else:
	    file_name_short=file_name
    
    fo = open(file_name,"r+")
    foo = open(file_name.rsplit("/",1)[0]+"/massmatrix.log","w+")
    lines =  fo.read().splitlines()
    toearly=0;
    for j in range(len(lines)):
    	if "0                MJJZ   " in lines[j]:
		toearly=1;
	if toearly==1:
		if "COLUMN" in lines[j]:
			if "ARE NULL" not in lines[j]:
				if "PAGE" not in lines[j+1]:
    					foo.write(lines[j+1]+"\n")
				else:
					foo.write(lines[j+7]+"\n")
				
    
    current_button=Tkinter.Button(frames[i], text="Selected file:"+file_name_short,state=DISABLED)
    current_button.grid(row=2, column=2, sticky='EW', padx=8, pady=2)
    
    
    
    
    

def outdirselection(i):

    """Returns an opened file in read mode."""
    dir_opt = options = {}
    options['initialdir'] = 'C:\\'
    options['mustexist'] = False
    options['parent'] = form
    options['title'] = 'This is a title'
    
    
    dir_name=tkFileDialog.askdirectory(**dir_opt)
    
    global config
    #global currentfile
    config.write("outdir: \n")
    config.write(dir_name)
    config.write("\n")
    
    if len(dir_name)>100:
	    dir_name_short="..."+dir_name[(len(dir_name)-100):len(dir_name)]
    else:
	    dir_name_short=dir_name

    
    current_button=Tkinter.Button(frames[i], text="Selected output directory:"+dir_name_short,state=DISABLED)
    current_button.grid(row=8, column=2, sticky='EW', padx=8, pady=2)

    


def startexe(i):	
	
	global config
	config.close()
	config_read = open("config.log","r")
	problems=0
	lines =  config_read.read().splitlines()
	for j in range(len(lines)):
    		if "File:0 :" in lines[j]:
        		firstfile=lines[j+1]		
			problems=problems+1
		elif "File:1 :" in lines[j]:
			secondfile=lines[j+1]
			problems=problems+1
		elif "outdir:" in lines[j]:
			outdir=lines[j+1]
			problems=problems+1
		
	if(problems<3):
		top = Toplevel()
		top.title("!FATAL WARNING!")
		top.minsize(width=450, height=266)
    		top.maxsize(width=450, height=266)
		msg = Message(top, text="You did not enter all the stuf bro!")
		msg.grid(row=1, column=2, sticky='EW', padx=12, pady=2)
		button = Button(top, text="Ok bro I check it out", command=top.destroy)
		button.grid(row=2, column=2, sticky='EW', padx=8, pady=2)
	else:
		if firstfile.endswith(".op2"):
			firsttype="NASTRAN"
		else:
			firsttype="ABAQUS"
		if secondfile.endswith(".op2"):
			secondtype="NASTRAN"
		else:
			secondtype="ABAQUS"
			
		os.system("\"C:\\Program Files (x86)\\BETA_CAE_Systems\\meta_post_v14.1.2\\meta_post64.bat"+"\""+" -s \\metacompare_windows.bs"" "+" -nogui "+firstfile+" "+firsttype+" "+secondfile+" "+secondtype+" "+outdir)	
		
		outlog=outdir+"\\metacompare.log"
		#os.system("rm -rf "+outlog)
		os.remove(outlog)
		while(os.path.isfile(outlog)==1):
			time.sleep(1.0)
		if(os.path.isfile(outlog)==0):
			output=open(outdir+"\\metaresult.temp")
			outlines =  output.read().splitlines()
			numberofrows=outlines[0]
			numberofcolumns=outlines[1]
			
			
			firstfilename=firstfile.split("\\")[-1]
			secondfilename=secondfile.split("\\")[-1]
			print firstfilename 
			print secondfilename
			print outdir+"\\Similarity.xlsx"
			current_workbook = xlsxwriter.Workbook(outdir+"\\Similarity.xlsx")
			current_worksheet=current_workbook.add_worksheet("Coaxiality_Table")
			
			merge_format_border = current_workbook.add_format({'bold':1,'border': 1})
			
			merge_format_border_green = current_workbook.add_format({'bold':1,'border': 1})
			merge_format_border_green.set_font_color('green')
			merge_format_border_orange = current_workbook.add_format({'bold':1,'border': 1})
			merge_format_border_orange.set_font_color('gray')
			merge_format_border_red = current_workbook.add_format({'bold':1,'border': 1})
			merge_format_border_red.set_font_color('red')
			
			
			merge_format_grau = current_workbook.add_format({
		        'bold': 1,
    			'border': 1,
    			'align': 'center',
    			'valign': 'vcenter',
    			'fg_color': '#E0E0E0'})
			
			
			merge_format_grau_vert = current_workbook.add_format({
		        'bold': 1,
    			'border': 1,
    			'align': 'center',
    			'fg_color': '#E0E0E0'})
			
			merge_format_grau_vert.set_rotation(90)
			
			start_row=3;
			start_column=3;
			first_eigen=[]
			second_eigen=[]
			current_worksheet.write(start_row,start_column,"Model1/Model2", merge_format_border)
			
			
			
			current_worksheet.write(0,0,"Model 1",merge_format_border)
			current_worksheet.write(1,0,"Model 2",merge_format_border)
			current_worksheet.write(0,1,firstfilename,merge_format_border)
			current_worksheet.write(1,1,secondfilename,merge_format_border)
			current_worksheet.write(0,2,firsttype,merge_format_border)
			current_worksheet.write(1,2,secondtype,merge_format_border)
			
			
			for j in range(2,int(numberofrows)+2):
				first_eigen.append(outlines[j])
				
			for j in range(2+int(numberofrows),int(numberofcolumns)+2+int(numberofrows)):
				second_eigen.append(outlines[j])
				
			for j in range(int(numberofrows)):
				current_worksheet.write(start_row+j+1,start_column,j+1,merge_format_border)
			for j in range(int(numberofcolumns)):
				current_worksheet.write(start_row,start_column+j+1,j+1,merge_format_border)
				
			#current_row=1
			#current_column=1				
			for j in range(2+int(numberofrows)+int(numberofcolumns),len(outlines)):
				current_row=start_row+int(outlines[j].split(",")[0])
				current_column=start_column+int(outlines[j].split(",")[1])
				current_value=float(outlines[j].split(",")[2])
				if int(outlines[j].split(",")[0])<7 and int(outlines[j].split(",")[1])<7:
					current_worksheet.write(current_row,current_column,current_value,merge_format_border_red)
				elif int(outlines[j].split(",")[0])>=7 and int(outlines[j].split(",")[1])>=7:
					current_worksheet.write(current_row,current_column,current_value,merge_format_border_green)
				else:
					current_worksheet.write(current_row,current_column,current_value,merge_format_border_orange)
			
			current_row=current_row+1
			current_worksheet.merge_range(current_row,start_column+1,current_row,start_column+6,'Rigidoelastic Columns',merge_format_grau)
			current_worksheet.merge_range(current_row,start_column+7,current_row,current_column,'Elastic Columns',merge_format_grau)
			start_row=4
			current_worksheet.merge_range(start_row,current_column+1,start_row+6,current_column+1,'Rigidoelastic Rows',merge_format_grau_vert)
			current_worksheet.merge_range(start_row+7,current_column+1,current_row-1,current_column+1,'Elastic Rows',merge_format_grau_vert)
			
			#start_row=3
			#current_worksheet.conditional_format(start_row+1,start_column+1,current_row,current_column, {'type': 'data_bar', 'criteria': '%'})
			
			current_row=3;
			start_row=current_row+1
			current_column=0;
			
			current_worksheet.write(current_row,current_column,"Eigenvalues M1", merge_format_border)
			current_column=current_column+1
			
			current_worksheet.write(current_row,current_column,"Eigenvalues M2", merge_format_border)
			current_column=current_column+1
			
			
			
			current_worksheet.write(current_row,current_column,"M1 % M2",merge_format_border)
			current_row=current_row+1
			current_column=current_column-2
			
			for j in range(int(numberofrows)):
				current_worksheet.write(current_row,current_column,float(first_eigen[j]),merge_format_border)
				current_row=current_row+1
			current_row=start_row
			current_column=current_column+1
			for j in range(int(numberofcolumns)):
				current_worksheet.write(current_row,current_column,float(second_eigen[j]),merge_format_border)
				current_row=current_row+1
			current_row=start_row
			current_column=current_column+1
			for j in range(6):
				current_worksheet.write(current_row,current_column,"*",merge_format_border)
				current_row=current_row+1
			for j in range(6,min(int(numberofrows),int(numberofcolumns))):
				current_worksheet.write(current_row,current_column,float(first_eigen[j])/float(second_eigen[j])*100.0,merge_format_border)
				current_row=current_row+1
				
				
			current_worksheet.set_column('A:C',40)
			current_worksheet.set_column('D:D',40)
			
			current_chart = current_workbook.add_chart({'type': 'line'})
			
			
			current_chart.add_series({'name':'=Coaxiality_Table!$A$4','categories':'=Coaxiality_Table!$D$11:$D$30','values':'=Coaxiality_Table!$A$11:$A$30'})
			current_chart.add_series({'name':'=Coaxiality_Table!$B$4','categories':'=Coaxiality_Table!$D$11:$D$30','values':'=Coaxiality_Table!$B$11:$B$30'})
			
			current_chart.set_title ({'name': 'Eigenvalue Comparison'})
			current_chart.set_x_axis({'name': 'Eigenvalue Number'})
			current_chart.set_y_axis({'name': 'Eigenvalue Value'})
			current_chart.set_style(10)
			current_chart.set_size({'width': 720, 'height': 720})
			current_worksheet.insert_chart('A31', current_chart, {'x_offset': 1, 'y_offset': 1})
			
			current_workbook.close()
			
			
			current_button=Tkinter.Button(frames[i], text="Processing is done check your output",state=DISABLED)
        		current_button.grid(row=10, column=2, sticky='EW', padx=8, pady=2)

	#print firstfile
	#print secondfile 
        #print outdir 
	
	#os.system("./Imager_MASTER.exe")
	

			
if __name__ == '__main__':
    columsize=25
    
    config = open("./config.log","w+")
    form = Tkinter.Tk()
    form.minsize(width=1200, height=666)
    form.maxsize(width=1200, height=666)

    #getFld = Tkinter.IntVar()

    form.wm_title('Configuration')


    n = notebook(form, LEFT)
    frames=[]
    buttons=[]
    entries=[]
    #start: the header section
    current_frame=Tkinter.Frame(n())
    frames.append(current_frame)
    current_frame=Tkinter.Frame(n())
    frames.append(current_frame)
    current_frame=Tkinter.Frame(n())
    frames.append(current_frame)
    current_frame=Tkinter.Frame(n())
    frames.append(current_frame)
    current_frame=Tkinter.Frame(n())
    frames.append(current_frame)

    current_button=Tkinter.Button(frames[0], text="Load Model 1 result",command=lambda:fileselection(0))
    buttons.append(current_button)
    buttons[len(buttons)-1].grid(row=1, column=2, sticky='EW', padx=8, pady=2)
   
    current_button=Tkinter.Button(frames[1], text="Load Model 2 Result",command=lambda:fileselection(1))
    buttons.append(current_button)
    buttons[len(buttons)-1].grid(row=1, column=2, sticky='EW', padx=8, pady=2)
   
    current_button=Tkinter.Button(frames[2], text="Load Model 1 Mass Matrix f06",command=lambda:f06selection(2))
    buttons.append(current_button)
    buttons[len(buttons)-1].grid(row=1, column=2, sticky='EW', padx=8, pady=2)

    current_button=Tkinter.Button(frames[3], text="Select Output Directory",command=lambda:outdirselection(3))
    buttons.append(current_button)
    buttons[len(buttons)-1].grid(row=1, column=2, sticky='EW', padx=8, pady=2)
     
    current_button=Tkinter.Button(frames[4], text="Submit for Comparison",command=lambda:startexe(4))
    buttons.append(current_button)
    buttons[len(buttons)-1].grid(row=1, column=2, sticky='EW', padx=8, pady=2)


    # keeps the reference to the radiobutton (optional)
    x1 = n.add_screen(frames[0], "Upload 1st Result")
    n.add_screen(frames[1], "Upload 2nd Result")
    n.add_screen(frames[2], "Upload MGG Result")
    n.add_screen(frames[3], "Select Output Directory")
    n.add_screen(frames[4], "Submit for Comparison")

    form.mainloop()
    exit()