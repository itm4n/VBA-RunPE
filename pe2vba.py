from subprocess import Popen, PIPE 
from os import system, remove
from os.path import isfile
import argparse

MAX_PROC_SIZE = 50 # Nbr of lines per procedure 
MAX_LINE_SIZE = 50 # Nbr of bytes per line

def is_printable(c):
  # All characters from SPACE to ~ are printable ASCII.
  # However, we want to avoid '"' 
  if c >= 0x20 and c < 0x7F and c != 0x22:
    return True
  else:
    return False

def pe_to_vba(pe):
  block = "" 
  line = ""
  ba = bytearray(pe)
  
  blocks = []
  cnt_bytes_current_line = 0
  cnt_lines_current_block = 0
  cnt_bytes_total = 0 
  prev_char_was_printable = False
  
  for b in ba:
  
    if cnt_lines_current_block == 0:
      # Start a new block 
      block = "    strPE = \"\"\n"   
      cnt_lines_current_block += 1 
    if cnt_bytes_current_line == 0:
      # Start a new line 
      line = "strPE"
    
    if is_printable(b):
      if prev_char_was_printable:
        line += chr(b)
      else:
        line = "B(%s, \"%s" % (line, chr(b))
      prev_char_was_printable = True
    else:
      if prev_char_was_printable:
        line += "\")"
      line = "A(%s, %s)" % (line, str(b))
      prev_char_was_printable = False
    
    cnt_bytes_current_line += 1 # We added a byte so increment the counter 
    cnt_bytes_total += 1 # We treated one more byte in the overall file 

    # If we reach the max number of bytes per line or the end of the array 
    # then end the current line.
    if cnt_bytes_current_line == MAX_LINE_SIZE or cnt_bytes_total == len(ba):
      if prev_char_was_printable:
        block += "    strPE = %s\")\n" % (line)
      else:
        block += "    strPE = %s\n" % (line)
      prev_char_was_printable = False # Must reset this flag for each new line 
      cnt_bytes_current_line = 0
      cnt_lines_current_block += 1

    # If we reach the max number of lines per block or the end of the array
    # then end the current block. 
    if cnt_lines_current_block == MAX_PROC_SIZE or cnt_bytes_total == len(ba):
      cnt_lines_current_block = 0 # Reset the number of lines 
      cnt_bytes_current_line = 0 # Reset the number of bytes per line 
      blocks.append(block) # Append the current block to the list of procedudes 
  
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


