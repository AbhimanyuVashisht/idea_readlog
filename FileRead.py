from datetime import datetime
import sys
# way to run : python FileRead.py <LOG file name> <IMSI>
# dt to name the file using time log as well
dt = str(datetime.now())
input_file = sys.argv[1]
to_trace = sys.argv[2]
infile_read = open(input_file, "r")
# Naming technique to name the filtered log file using IMSI
filename = to_trace + "_" + dt + "_" + ".LOG"

infile_write = open(filename, "a")
record_count = 0
trace_found = False
MSCi_data = []
for line in infile_read:
    if "MSCi" in line:
        for i in range(0, 6, 1):
            MSCi_data.append(line)
            line = next(infile_read)

        MSCi_data.append(line)

        line_split = line.split(':')
        # checking the traced IMSi
        if line_split[1].strip() == sys.argv[2]:
            trace_found = True
            record_count = record_count + 1
            for i in range(0, 7, 1):
                infile_write.write(MSCi_data[i])

            MSCi_data = []

        else:
            trace_found = False
            MSCi_data = []

    elif trace_found and "END OF REPORT" not in line:
        infile_write.write(line)
    elif trace_found and "END OF REPORT" in line:
        infile_write.write("END OF REPORT\r\n\r\n\r\n\r\n\r\n")
        trace_found = False
    elif "END OF REPORT":
        trace_found = False

