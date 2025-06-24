from selenium.webdriver.support.ui import WebDriverWait

import time

def musteri_yontemi(df,driver,By,wait,EC,ActionChains):
    for index,row in df.iterrows():
        # if index == 10:
        #     break
        add_row_button = driver.find_element(By.ID, "grdLedgerVoucherxGrid_an_0")
        add_row_button.click()

        if row["İşlem Grubu"] == "Avans ve Kira":
            select_input = driver.find_element(By.ID, f"grdLedgerVoucherxGrid_rc_{index}_1")
            select_input.click()

        islem_grubu_block = driver.find_element(By.ID, f"grdLedgerVoucherxGrid_rc_{index}_2").find_element(By.CSS_SELECTOR, "nobr")
        islem_grubu_block.click()
        islem_grubu_input = driver.find_element(By.ID, "grdLedgerVoucherxGrid_tb")

        if row["İşlem Grubu"] == "Avans ve Kira":
            islem_grubu_input.send_keys("kira")
        else:
            islem_grubu_input.send_keys(row["İşlem Grubu"])

        #time.sleep(0.5)
        #islem_grubu_option = driver.find_element(By.ID, "Tr0")
        islem_grubu_option = wait.until(EC.element_to_be_clickable((By.ID, "Tr0")))
        islem_grubu_option.click()

        #####müşteri block#####
        musteri_block = driver.find_element(By.ID, f"grdLedgerVoucherxGrid_rc_{index}_7")
        main_window = driver.current_window_handle
        musteri_block.click()

        # go to menu window
        wait.until(EC.number_of_windows_to_be(2))
        for window_handle in driver.window_handles:
            if window_handle != main_window:
                driver.switch_to.window(window_handle)
                break

        musteri_karti_block = driver.find_element(By.ID, "WebGridSearch_rc_0_1").find_element(By.CSS_SELECTOR, "nobr")
        actions = ActionChains(driver)
        actions.double_click(musteri_karti_block).perform()

        musteri_karti_input = driver.find_element(By.ID, "WebGridSearch_tb")
        musteri_karti_input.send_keys(row["hesap kartno"].split('.')[4])

        search_button = driver.find_element(By.CLASS_NAME, "ig_38e2bc54_r1")
        search_button.click()
        #time.sleep(1)

        try:
            #searched_row = driver.find_element(By.ID, "WebGrid_l_0")
            searched_row = wait.until(EC.element_to_be_clickable((By.ID, "WebGrid_l_0")))
            searched_row.click()
        except:
            pass
        
        ok_button = driver.find_element(By.ID, "btnOk")
        ok_button.click()
        
        # back to main window
        if len(driver.window_handles) > 1:
            driver.close()

        wait.until(EC.number_of_windows_to_be(1))
        for window_handle in driver.window_handles:
            if window_handle == main_window:
                driver.switch_to.window(window_handle)
                break
        
        driver.switch_to.default_content()
        #driver.switch_to.frame("main")
        WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it("main"))

        #####kira planı block###
        kira_block = driver.find_element(By.ID, f"grdLedgerVoucherxGrid_rc_{index}_10")
        main_window = driver.current_window_handle
        kira_block.click()

        # go to menu window
        wait.until(EC.number_of_windows_to_be(2))
        for window_handle in driver.window_handles:
            if window_handle != main_window:
                driver.switch_to.window(window_handle)
                break

        kira_karti_block = driver.find_element(By.ID, "WebGridSearch_rc_0_1").find_element(By.CSS_SELECTOR, "nobr")
        actions = ActionChains(driver)
        actions.double_click(kira_karti_block).perform()

        kira_karti_input = driver.find_element(By.ID, "WebGridSearch_tb")
        kira_karti_input.send_keys(str(int(row["hesap kartno"].split('.')[5])))

        search_button = driver.find_element(By.CLASS_NAME, "ig_38e2bc54_r1")
        search_button.click()
        #time.sleep(1)

        try:
            #searched_row = driver.find_element(By.ID, "WebGrid_l_0")
            searched_row = wait.until(EC.element_to_be_clickable((By.ID, "WebGrid_l_0")))
            searched_row.click()
        except:
            pass
        
        ok_button = driver.find_element(By.ID, "btnOk")
        ok_button.click()
        
        # back to main window
        if len(driver.window_handles) > 1:
            driver.close()

        wait.until(EC.number_of_windows_to_be(1))
        for window_handle in driver.window_handles:
            if window_handle == main_window:
                driver.switch_to.window(window_handle)
                break
        
        driver.switch_to.default_content()
        #driver.switch_to.frame("main")
        WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it("main"))

        # kira_plani_block = wait.until(EC.element_to_be_clickable((By.ID, f"grdLedgerVoucherxGrid_rc_{index}_8"))).find_element(By.CSS_SELECTOR, "nobr")
        # kira_plani_block.click()
        # kira_plani_input = driver.find_element(By.ID, "grdLedgerVoucherxGrid_tb")
        # kira_plani_input.send_keys(str(int(row["hesap kartno"].split('.')[5])))
        # #time.sleep(0.3)
        # try:
        #     #kira_plani_option = driver.find_element(By.ID, f"Tr{str(int(row["hesap kartno"].split('.')[6])-1)}")
        #     kira_plani_option = wait.until(EC.element_to_be_clickable((By.ID, f"Tr{str(int(row["hesap kartno"].split('.')[6])-1)}")))
        #     kira_plani_option.click()
        # except:
        #     continue
        
        aciklama_block = driver.find_element(By.ID, f"grdLedgerVoucherxGrid_rc_{index}_17").find_element(By.CSS_SELECTOR, "nobr")
        aciklama_block.click()
        aciklama_input = driver.find_element(By.ID, "grdLedgerVoucherxGrid_tb")
        aciklama_input.send_keys(str(row["Açıklama"]))

        b_a_block = driver.find_element(By.ID, f"grdLedgerVoucherxGrid_rc_{index}_18").find_element(By.CSS_SELECTOR, "nobr")
        b_a_block.click()
        b_a_input = driver.find_element(By.ID, "grdLedgerVoucherxGrid_tb")
        b_a_input.send_keys(str(row["B/A"]))

        tutar_block = driver.find_element(By.ID, f"grdLedgerVoucherxGrid_rc_{index}_19").find_element(By.CSS_SELECTOR, "nobr")
        tutar_block.click()
        tutar_input = driver.find_element(By.ID, "igtxtgrdLedgerVoucher_DoubleColumn")
        tutar_input.send_keys(str(row["Tutar"]).replace(".",","))

        # kur_block = driver.find_element(By.ID, f"grdLedgerVoucherxGrid_rc_{index+1}_20").find_element(By.CSS_SELECTOR, "nobr")
        # kur_block.click()
        # kur_input = driver.find_element(By.ID, "igtxtgrdLedgerVoucher_DoubleColumnCustom")
        # kur_input.send_keys("1")

def banka_yontemi(df,driver,By,wait,EC,ActionChains):
    ########more card 1
    add_row_button = driver.find_element(By.ID, "grdLedgerVoucherxGrid_an_0")
    add_row_button.click()

    islem_grubu_block = driver.find_element(By.ID, f"grdLedgerVoucherxGrid_rc_0_2").find_element(By.CSS_SELECTOR, "nobr")
    islem_grubu_block.click()
    islem_grubu_input = driver.find_element(By.ID, "grdLedgerVoucherxGrid_tb")
    islem_grubu_input.send_keys("banka")

    islem_grubu_option = wait.until(EC.element_to_be_clickable((By.ID, "Tr0")))
    islem_grubu_option.click()

    banka_block = driver.find_element(By.ID, f"grdLedgerVoucherxGrid_rc_0_7")
    main_window = driver.current_window_handle
    banka_block.click()

    # go to menu window
    wait.until(EC.number_of_windows_to_be(2))
    for window_handle in driver.window_handles:
        if window_handle != main_window:
            driver.switch_to.window(window_handle)
            break

    banka_karti_block = driver.find_element(By.ID, "WebGridSearch_rc_0_1").find_element(By.CSS_SELECTOR, "nobr")
    actions = ActionChains(driver)
    actions.double_click(banka_karti_block).perform()

    banka_karti_input = driver.find_element(By.ID, "WebGridSearch_tb")
    banka_karti_input.send_keys("b000401")

    search_button = driver.find_element(By.CLASS_NAME, "ig_38e2bc54_r1")
    search_button.click()

    try:
        searched_row = wait.until(EC.element_to_be_clickable((By.ID, "WebGrid_l_0")))
        searched_row.click()
    except:
        pass
    
    ok_button = driver.find_element(By.ID, "btnOk")
    ok_button.click()
    
    # back to main window
    if len(driver.window_handles) > 1:
        driver.close()

    wait.until(EC.number_of_windows_to_be(1))
    for window_handle in driver.window_handles:
        if window_handle == main_window:
            driver.switch_to.window(window_handle)
            break
    
    driver.switch_to.default_content()
    WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it("main"))

    b_a_block = wait.until(EC.element_to_be_clickable((By.ID, "grdLedgerVoucherxGrid_rc_0_15"))).find_element(By.CSS_SELECTOR, "nobr")
    b_a_block.click()
    b_a_input = driver.find_element(By.ID, "grdLedgerVoucherxGrid_tb")
    b_a_input.send_keys("B")

    aciklama_block = wait.until(EC.element_to_be_clickable((By.ID, "grdLedgerVoucherxGrid_rc_0_16"))).find_element(By.CSS_SELECTOR, "nobr")
    aciklama_block.click()
    aciklama_input = driver.find_element(By.ID, "grdLedgerVoucherxGrid_tb")
    aciklama_input.send_keys("Tapu Devri Nedeniyle Cari Virmanı")

    tutar_block = driver.find_element(By.ID, f"grdLedgerVoucherxGrid_rc_0_17").find_element(By.CSS_SELECTOR, "nobr")
    tutar_block.click()
    tutar_input = driver.find_element(By.ID, "igtxtgrdLedgerVoucher_DoubleColumn")
    tutar_input.send_keys("1")

    ######more card 2
    add_row_button = driver.find_element(By.ID, "grdLedgerVoucherxGrid_an_0")
    add_row_button.click()

    islem_grubu_block = driver.find_element(By.ID, f"grdLedgerVoucherxGrid_rc_1_2").find_element(By.CSS_SELECTOR, "nobr")
    islem_grubu_block.click()
    islem_grubu_input = driver.find_element(By.ID, "grdLedgerVoucherxGrid_tb")
    islem_grubu_input.send_keys("banka")

    islem_grubu_option = wait.until(EC.element_to_be_clickable((By.ID, "Tr0")))
    islem_grubu_option.click()

    banka_block = driver.find_element(By.ID, f"grdLedgerVoucherxGrid_rc_1_7")
    main_window = driver.current_window_handle
    banka_block.click()

    # go to menu window
    wait.until(EC.number_of_windows_to_be(2))
    for window_handle in driver.window_handles:
        if window_handle != main_window:
            driver.switch_to.window(window_handle)
            break

    banka_karti_block = driver.find_element(By.ID, "WebGridSearch_rc_0_1").find_element(By.CSS_SELECTOR, "nobr")
    actions = ActionChains(driver)
    actions.double_click(banka_karti_block).perform()

    banka_karti_input = driver.find_element(By.ID, "WebGridSearch_tb")
    banka_karti_input.send_keys("b000401")

    search_button = driver.find_element(By.CLASS_NAME, "ig_38e2bc54_r1")
    search_button.click()

    try:
        searched_row = wait.until(EC.element_to_be_clickable((By.ID, "WebGrid_l_0")))
        searched_row.click()
    except:
        pass
    
    ok_button = driver.find_element(By.ID, "btnOk")
    ok_button.click()
    
    # back to main window
    if len(driver.window_handles) > 1:
        driver.close()

    wait.until(EC.number_of_windows_to_be(1))
    for window_handle in driver.window_handles:
        if window_handle == main_window:
            driver.switch_to.window(window_handle)
            break
    
    driver.switch_to.default_content()
    WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it("main"))

    b_a_block = wait.until(EC.element_to_be_clickable((By.ID, f"grdLedgerVoucherxGrid_rc_1_15"))).find_element(By.CSS_SELECTOR, "nobr")
    b_a_block.click()
    b_a_input = driver.find_element(By.ID, "grdLedgerVoucherxGrid_tb")
    b_a_input.send_keys("A")

    aciklama_block = driver.find_element(By.ID, f"grdLedgerVoucherxGrid_rc_1_16").find_element(By.CSS_SELECTOR, "nobr")
    aciklama_block.click()
    aciklama_input = driver.find_element(By.ID, "grdLedgerVoucherxGrid_tb")
    aciklama_input.send_keys("Tapu Devri Nedeniyle Cari Virmanı")

    tutar_block = driver.find_element(By.ID, f"grdLedgerVoucherxGrid_rc_1_17").find_element(By.CSS_SELECTOR, "nobr")
    tutar_block.click()
    tutar_input = driver.find_element(By.ID, "igtxtgrdLedgerVoucher_DoubleColumn")
    tutar_input.send_keys("1")

    for index,row in df.iterrows():
        add_row_button = driver.find_element(By.ID, "grdLedgerVoucherxGrid_an_0")
        add_row_button.click()

        select_input = driver.find_element(By.ID, f"grdLedgerVoucherxGrid_rc_{index+2}_0")
        select_input.click()

        pb_block = driver.find_element(By.ID, f"grdLedgerVoucherxGrid_rc_{index+2}_8").find_element(By.CSS_SELECTOR, "nobr")
        pb_block.click()
        pb_input = driver.find_element(By.ID, "grdLedgerVoucherxGrid_tb")
        pb_input.send_keys("tl")
        pb_option = driver.find_element(By.ID, "Tr0")
        pb_option.click()

        time.sleep(0.5)

        hesap_block = driver.find_element(By.ID, f"grdLedgerVoucherxGrid_rc_{index+2}_13")
        main_window = driver.current_window_handle
        hesap_block.click()

        # go to menu window
        wait.until(EC.number_of_windows_to_be(2))
        for window_handle in driver.window_handles:
            if window_handle != main_window:
                driver.switch_to.window(window_handle)
                break

        musteri_karti_block = driver.find_element(By.ID, "WebGridSearch_rc_0_1").find_element(By.CSS_SELECTOR, "nobr")
        actions = ActionChains(driver)
        actions.double_click(musteri_karti_block).perform()

        musteri_karti_input = driver.find_element(By.ID, "WebGridSearch_tb")
        musteri_karti_input.send_keys(row["hesap kartno"])

        search_button = driver.find_element(By.CLASS_NAME, "ig_38e2bc54_r1")
        search_button.click()
        time.sleep(1)

        try:
            searched_row = driver.find_element(By.ID, "WebGrid_l_0")
            searched_row.click()
        except:
            pass
        
        ok_button = driver.find_element(By.ID, "btnOk")
        ok_button.click()
        
        # back to main window
        if len(driver.window_handles) > 1:
            driver.close()

        wait.until(EC.number_of_windows_to_be(1))
        for window_handle in driver.window_handles:
            if window_handle == main_window:
                driver.switch_to.window(window_handle)
                break
        
        driver.switch_to.default_content()
        #driver.switch_to.frame("main")
        WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it("main"))

        b_a_block = wait.until(EC.element_to_be_clickable((By.ID, f"grdLedgerVoucherxGrid_rc_{index+2}_15"))).find_element(By.CSS_SELECTOR, "nobr")
        b_a_block.click()
        b_a_input = driver.find_element(By.ID, "grdLedgerVoucherxGrid_tb")
        b_a_input.send_keys(str(row["B/A"]))
        
        aciklama_block = driver.find_element(By.ID, f"grdLedgerVoucherxGrid_rc_{index+2}_16").find_element(By.CSS_SELECTOR, "nobr")
        aciklama_block.click()
        aciklama_input = driver.find_element(By.ID, "grdLedgerVoucherxGrid_tb")
        aciklama_input.send_keys(str(row["Açıklama"]))

        tutar_block = driver.find_element(By.ID, f"grdLedgerVoucherxGrid_rc_{index+2}_17").find_element(By.CSS_SELECTOR, "nobr")
        tutar_block.click()
        tutar_input = driver.find_element(By.ID, "igtxtgrdLedgerVoucher_DoubleColumn")
        tutar_input.send_keys(str(row["Tutar"]).replace(".",","))

        # kur_block = driver.find_element(By.ID, f"grdLedgerVoucherxGrid_rc_{index}_20").find_element(By.CSS_SELECTOR, "nobr")
        # kur_block.click()
        # kur_input = driver.find_element(By.ID, "igtxtgrdLedgerVoucher_DoubleColumnCustom")
        # kur_input.send_keys("1")
