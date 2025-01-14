import numpy as np

def get_numberstring_two_dig(value):
  if isinstance(value, (float)):
      value = str(int(value))
  elif np.isnan(value):
      value = ''
      
  if(value == 0):
    value = "00"
  else:
      value = str(value)
      
  if(len(value) <= 1):
    value = "0" + value
    
  return value

def get_numberstring_three_dig(value):
    if isinstance(value, (int, float)):
        if np.isnan(value):
            value = ''
        else:
            value = str(int(value))  
            
    elif isinstance(value, str):
        value = str(value) 

    if value == "0":
        return "000"

    if len(value) <= 1:
        value = "00" + value
    elif len(value) == 2:
        value = "0" + value
    
    return value