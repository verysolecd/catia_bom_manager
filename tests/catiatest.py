def recurPrd(self, oPrd, LV, all_rows_data=None, counter=1):
    if all_rows_data is None:
        all_rows_data = []
    row_data = [counter, LV, *self.info_Prd(oPrd)]
    all_rows_data.append(row_data)
    if oPrd.products.count > 0:
        for i in range(1, oPrd.products.count + 1):
            child = oPrd.products.item(i)
            counter += 1
            all_rows_data = self.recurPrd(
                child, LV + 1, all_rows_data, counter)
    return all_rows_data

    # if oPrd.products.count > 0:
    #     for i in range(1, oPrd.products.count + 1):
    #         child = oPrd.products.item(i)
    #         self._collect_bom_rows(child, LV + 1, all_rows_data)
