def combine_people(arrays):
    result = []
    for data in arrays:
        tuples_list = [tuple(data[i:i+4]) for i in range(0, len(data), 4)]
        unique_ids = {}
        for tup in tuples_list:
            if tup[1] not in unique_ids:
                unique_ids[tup[1]] = tup
        unique_data = [elem for unique_id, elem in unique_ids.items()]
        unique_array = [item for sublist in unique_data for item in sublist]
        result.append(unique_array)
    return result



