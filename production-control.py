# Import Modules
from tkinter import *
from pc_pkg import *

# Create Application Window
root = Tk()
root.title('HJ Reporting')
root.geometry('370x555')


# Grid File Details
@cache
def gridDP():
    global stats

    verification = findFileDP() + '\n' + statsDP()
    v_label = Label(dp_frame, text=verification)
    v_label.grid(row=4, column=0)


@cache
def gridFP():
    global stats

    verification = findFileFP() + '\n' + statsFP()
    v_label = Label(fp_frame, text=verification)
    v_label.grid(row=4, column=0)


# Frame Reports
dp_frame = LabelFrame(root, text='Production Control DP', font=("Arial Bold", 9))
fp_frame = LabelFrame(root, text='Production Control FP', font=("Arial Bold", 9))

# Grid Reports
dp_frame.grid(row=0, column=0, pady=5, padx=5)
fp_frame.grid(row=1, column=0, pady=5, padx=5)

# Clear Report Buttons
Button(dp_frame, text="Clear DP Report", command=clearDP).grid(row=0, column=0, columnspan=2, pady=5, padx=10,
                                                               ipadx=121)
Button(fp_frame, text="Clear FP Report", command=clearFP).grid(row=0, column=0, columnspan=2, pady=5, padx=10,
                                                               ipadx=121)

# Pull Report Buttons
Button(dp_frame, text="Pull Report", command=gridDP).grid(row=1, column=0, columnspan=2, pady=5, padx=10, ipadx=134)
Button(fp_frame, text="Pull Report", command=gridFP).grid(row=1, column=0, columnspan=2, pady=5, padx=10, ipadx=134)

# Verify & Run Buttons
Button(dp_frame, text="Verify & Run Report", fg='green', command=reportDP).grid(row=2, column=0, columnspan=2, pady=5,
                                                                                padx=10,
                                                                                ipadx=110)
Button(fp_frame, text="Verify & Run Report", fg='green', command=reportFP).grid(row=2, column=0, columnspan=2, pady=5,
                                                                                padx=10,
                                                                                ipadx=110)

# Verify Labels
Label(dp_frame, text="Report Details: ").grid(row=3, column=0)
Label(fp_frame, text="Report Details: ").grid(row=3, column=0)

# Check For Errors Buttons
Button(dp_frame, text="Check For Errors", fg='red', command=checkDP).grid(row=5, column=0, columnspan=2, pady=5,
                                                                          padx=10,
                                                                          ipadx=120)
Button(fp_frame, text="Check For Errors", fg='red', command=checkFP).grid(row=5, column=0, columnspan=2, pady=5,
                                                                          padx=10,
                                                                          ipadx=120)

# Run Application
root.mainloop()
