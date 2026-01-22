def page_numbers_to_hyphens(page_nums:list):
    """
    converts a list of page numbers into a string with hyphens for consecutive numbers
    """

    if not page_nums:
        return ""


    result_string = ""       
    page_nums = sorted(set(page_nums))
    previous_num = page_nums[0]
    tmp_list =[previous_num]

    for i, page_num in enumerate(page_nums[1:]):
        if not isinstance(page_num, int) or page_num < 1:
            raise ValueError("Page numbers must be positive integers.")

        if len(page_nums) == 1:
            return str(page_nums[0])

        elif len(page_nums) == 2:
            return f"{page_nums[0]}, {page_nums[1]}"

        elif page_num == previous_num + 1:
            previous_num = page_num
            tmp_list.append(page_num)

            # if last page number in the list
            if i == len(page_nums)-2:
                if len(tmp_list) > 1:
                    result_string += f"{tmp_list[0]}-{tmp_list[-1]}"
                else:
                    result_string += f"{page_num}"
    
        else:
            if len(tmp_list) > 1:
                result_string += f"{tmp_list[0]}-{tmp_list[-1]}, {page_num},"
            else:
                result_string += f"{page_num}, "

                # Remove the last comma for the last element
                if i == len(page_nums)-2:
                    result_string = result_string.rstrip(", ")
            previous_num = page_num
            tmp_list = []

    return result_string

if __name__ == '__main__':
    pages = [22,41]
    result = page_numbers_to_hyphens(pages)
    print(result)