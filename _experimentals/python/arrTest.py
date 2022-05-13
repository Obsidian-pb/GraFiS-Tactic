#%%
import numpy as np
import matplotlib.pyplot as plt


# %%
arr = np.loadtxt("ToMatrixTest.csv", delimiter=",")
print(arr.shape)
# plt.imshow(arr)

#%%

yx_0 = [[125,50]]
arr[yx_0[0][0],yx_0[0][1]] = 2

plt.imshow(arr)



# %%
for i in range(30):
    for point in yx_0:
        y,x = point[0],point[1]
        yx_0.remove(point)
        
        xn,yn=x-1,y
        if arr[yn,xn]==0:
            arr[yn,xn]=2
            yx_0.append([yn,xn])
        xn,yn=x-1,y+1
        if arr[xn,yn]==0:
            arr[yn,xn]=2
            yx_0.append([yn,xn])  
        xn,yn=x,y+1
        if arr[xn,yn]==0:
            arr[yn,xn]=2
            yx_0.append([yn,xn])
        xn,yn=x+1,y+1
        if arr[xn,yn]==0:
            arr[yn,xn]=2
            yx_0.append([yn,xn])
        xn,yn=x+1,y
        if arr[xn,yn]==0:
            arr[yn,xn]=2
            yx_0.append([yn,xn])
        xn,yn=x+1,y-1
        if arr[xn,yn]==0:
            arr[yn,xn]=2
            yx_0.append([yn,xn])
        xn,yn=x,y-1
        if arr[xn,yn]==0:
            arr[yn,xn]=2
            yx_0.append([yn,xn])
        xn,yn=x-1,y-1
        if arr[xn,yn]==0:
            arr[yn,xn]=2
            yx_0.append([yn,xn])           
             

plt.imshow(arr)



#%%
a=[1,2,3,4,5]
a.remove(4)
a