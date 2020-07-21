def subsetsum(array, num):
    if sum(array) == num:
        return array
    if len(array) > 1:
        for subset in (array[:-1], array[1:]):
            result = subsetsum(subset, num)
            if result is not None:
                return result


ar1 = "6	0	0	0	6	0	3	11	7	22	5	94	0	12	0	0	0	15	0	23	66	0	0	21	68	17	136	6	25	3	7	0	0"
ar2 = [int(num_str) for num_str in ar1.split('\t')]
ar = list(filter(lambda a: a != 0, ar2))

print(subsetsum(ar, 416))
input('Done!')