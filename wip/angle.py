from math import acos
from math import sqrt
from math import pi

ax = 10
ay = 10
bx = 15
by = 10
cx = 0
cy = 11

v=[ax-cx,ay-cy]
w=[bx-ax,by-ay]

print v
print w
  
cosx=(v[0]*w[0]+v[1]*w[1])/(sqrt(v[0]**2+v[1]**2)*sqrt(w[0]**2+w[1]**2))
rad=acos(cosx)
inner=rad*180/pi # returns degree

det = v[0]*w[1]-v[1]*w[0]

print det

if det<0: #this is a property of the det. If the det < 0 then B is clockwise of A
    print inner
else: # if the det > 0 then A is immediately clockwise of B
    print 360-inner