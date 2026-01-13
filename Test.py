print("program started")

sum = 0
for i in range(5):
    print(f"Iteration {i+1}")
    if i == 2:
        sum = sum + 2 * i

print ("Final sum:", sum)
print("program ended")