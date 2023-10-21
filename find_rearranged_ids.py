def filter_arrays(array_of_arrays):
    processed_ids = set()
    filtered_arrays = []

    for sub_array in array_of_arrays:
        ids = sub_array[1::4]
        sorted_ids = sorted(ids)
        sorted_ids_tuple = tuple(sorted_ids)

        if sorted_ids_tuple not in processed_ids:
            processed_ids.add(sorted_ids_tuple)
            filtered_arrays.append(sub_array)

    return filtered_arrays

