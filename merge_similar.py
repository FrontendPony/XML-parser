def merge_similar(arr):
    merged = []
    used_indices = set()

    for i, sub_arr in enumerate(arr):
        if i in used_indices:
            continue

        merged_sub = sub_arr[:]
        ids_to_match = {sub_arr[0], sub_arr[4]}

        for j, other_sub_arr in enumerate(arr[i + 1:], start=i + 1):
            if any(id_ in ids_to_match for id_ in (other_sub_arr[0], other_sub_arr[4])):
                used_indices.add(j)
                merged_sub = list((merged_sub + other_sub_arr))
                ids_to_match.update((other_sub_arr[0], other_sub_arr[4]))

        merged.append(merged_sub)

    return merged

