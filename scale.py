#scaling

def scale(value):
    if value < 44.1:
        value = value*(59/44)
    elif value < 59:
        value = (value-45)*(13/15)+60
    elif value < 74:
        value = (value-60)*(2/3)+73
    elif value >= 74:
        value = (value-75)*(18/26)+83
    if value > 100:
        value = 100
    return(value)
