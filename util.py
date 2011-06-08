#

#stolen
#http://code.activestate.com/recipes/142812/
FILTER=''.join([(len(repr(chr(x)))==3) and chr(x) or '.' for x in range(256)])

def hexdump(src, length=16, indent=0, addr=0):
    N=addr; result=''
    while src:
       s,src = src[:length],src[length:]
       hexa = ' '.join(["%02X"%ord(x) for x in s])
       s = s.translate(FILTER)
       result += (" "*indent) + "%010X   %-*s   %s\n" % (N, length*3, hexa, s)
       N+=length
    return result
