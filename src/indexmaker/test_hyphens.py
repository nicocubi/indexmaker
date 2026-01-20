def page_numbers_to_hyphens(pages:list):
    """
    converts a list of page numbers into a string with hyphens for consecutive numbers
    """

    if not pages:
        return ""


    result_string = ""       
    pages = sorted(set(pages))
    previous_num = pages[0]
    tmp_list =[previous_num]

    for i, page_num in enumerate(pages[1:]):
        print("page_num", page_num)
        if not isinstance(page_num, int) or page_num < 1:
            raise ValueError("Page numbers must be positive integers.")

        if page_num == previous_num + 1:
            previous_num = page_num
            tmp_list.append(page_num)
            if i == len(pages)-1:
                if len(tmp_list) > 1:
                    result_string += f"{tmp_list[0]}-{tmp_list[-1]}"
                else:
                    result_string += f"{page_num}"
    
        else:
            print("tmp_list", tmp_list)
            print("previous_num", previous_num)
            if len(tmp_list) > 1:
                result_string += f"{tmp_list[0]}-{tmp_list[-1]}, {page_num},"
            else:
                result_string += f"{page_num}, "
            previous_num = page_num
            tmp_list = []

    return result_string

if __name__ == '__main__':
    pages = (1,2,3,4,5,7,15,16)
    result = page_numbers_to_hyphens(pages)
    print(result)