from subprocess import Popen, PIPE 
from os import system, remove
from os.path import isfile
import argparse

MAX_PROC_SIZE = 100 # Nbr of lines per procedure 
MAX_LINE_SIZE = 10 # Nbr of bytes per lines 

def pe_to_vba(pe):
  result = "" 
  ba = bytearray(pe)
  
  blocks = []
  cnt_bytes = 0
  cnt_lines = 0 
  
  for b in ba:
  
    if cnt_lines == 0:
      # Start a new block 
      result = "    strPE = \"\"\n"   
      cnt_lines += 1 
    if cnt_bytes == 0:
      # Start a new line 
      result += "    strPE = strPE "
    
    result += " + " + "chr(&h" + format(b, "02x") + ")"  
    cnt_bytes += 1 # We added a byte so increment the counter 

    if cnt_bytes == MAX_LINE_SIZE:
      # If we reach the end of a line, add a LINE FEED 
      result += "\n"
      cnt_bytes = 0
      cnt_lines += 1
    
    if cnt_lines == MAX_PROC_SIZE:
      # If we reach max procedure size 
      cnt_lines = 0 # Reset the number of lines 
      cnt_bytes = 0 # Reset the number of bytes per line 
      blocks.append(result) # Append the current block to the list of procedudes 
  
  # Create a 'Sub' for each block  
  proc = ""
  for i in range(len(blocks)):
    proc += "Private Function PE" + str(i) + "() As String\n"
    proc += "   Dim strPE As String\n\n"
    proc += blocks[i]
    proc += "\n    PE" + str(i) + " = strPE\n"
    proc += "End Function\n\n"
  
  vba = ""
  vba += proc
  vba += "Private Function PE() As String\n"
  vba += "    Dim strPE As String\n"
  vba += "    strPE = \"\"\n"
  for i in range(len(blocks)):
    vba += "    strPE = strPE + PE" + str(i) + "()\n"
  vba += "    PE = strPE\n" 
  vba += "End Function\n" 
  
  return vba 

def main():

  # Parse command line arguments and options
  parser = argparse.ArgumentParser(description="PE to VBA file converter")
  parser.add_argument("pe_file", help="PE file to convert.")
  args = parser.parse_args()
  
  # Check whether input file exist 
  if not isfile(args.pe_file):
    print "[!] '%s' doesn't exist!" % (args.pe_file) 
    return 

  # Read the file 
  pe_file = open(args.pe_file, "rb") 
  pe = pe_file.read() 
  pe_file.close() 
  
  # Convert the file to VBA and write to file 
  out_filename = "%s.vba" % (args.pe_file)
  out_file = open(out_filename , "w") 
  out_file.write(pe_to_vba(pe)) 
  out_file.close()
  
  if isfile(out_filename): 
    print "[+] Created file '%s'." % (out_filename)
  
  return 

if __name__ == '__main__':
  main() 





