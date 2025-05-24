import requests

default_header = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0.0.0 Safari/537.36"
}

header = {
            "referer": "https://www.nseindia.com/",
             "Connection": "keep-alive",
             "Cache-Control": "max-age=0",
             "DNT": "1",
             "Upgrade-Insecure-Requests": "1",
             "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36",
             "Sec-Fetch-User": "?1",
             "Accept": "ext/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
             "Sec-Fetch-Site": "none",
             "Sec-Fetch-Mode": "navigate",
             "Accept-Language": "en-US,en;q=0.9,hi;q=0.8"
            }

def nse_urlfetch(url, origin_url="http://nseindia.com"):
    r_session = requests.session()
    nse_live = r_session.get(origin_url, headers=default_header)
    cookies = nse_live.cookies
    return r_session.get(url, headers=header, cookies=cookies)

def nse_eq_option(eq, origin_url="https://www.nseindia.com/option-chain"):
    return nse_urlfetch(f"https://www.nseindia.com/api/option-chain-equities?symbol={eq}",origin_url=origin_url).json()

def nse_indc_option(indc, origin_url="https://www.nseindia.com/option-chain"):
    return nse_urlfetch(f"https://www.nseindia.com/api/option-chain-indices?symbol={indc}",origin_url=origin_url).json()

indc="NIFTY"
eq="SBIN"
#print(nse_indc_option(indc))

# print(nse_eq_option("WIPRO")['records']['data'][0]['PE'].keys())
# print(nse_eq_option("WIPRO")['records']['data'])

# print(nse_urlfetch(f"https://www.nseindia.com/api/option-chain-equities?symbol={eq}","https://www.nseindia.com/option-chain").json()['records']['data'][0]['PE'])

# print(nse_urlfetch(f"https://www.nseindia.com/api/option-chain-indices?symbol={indc}","https://www.nseindia.com/option-chain").json()['records']['strikePrices'])

def eq_opt_ltp(eq, stk, exp, typ):
    option_data = nse_eq_option(eq)['records']['data']

    for item in option_data:
        strike = item.get('strikePrice')
        expiry = item.get('expiryDate')
        
        # print(stk,strike)
        # print(exp,expiry)
        if stk == strike and expiry == exp:
            if typ == 'CE':
                uv = item.get('CE', {}).get('underlyingValue')
                ltp = item.get('CE', {}).get('lastPrice')
                iv = item.get('CE', {}).get('impliedVolatility')
                oi = item.get('CE', {}).get('openInterest')
                coi = item.get('CE', {}).get('changeinOpenInterest')
                ttv = item.get('CE', {}).get('totalTradedVolume')               
                # print(f"Strike Price: {strike}, Expiry: {expiry}, Call LTP: {ltp}")
                return [eq, uv, stk, exp, typ, ltp, iv, oi, coi, ttv]
            elif typ == 'PE':
                uv = item.get('PE', {}).get('underlyingValue')
                ltp = item.get('PE', {}).get('lastPrice')
                iv = item.get('PE', {}).get('impliedVolatility')
                oi = item.get('PE', {}).get('openInterest')
                coi = item.get('PE', {}).get('changeinOpenInterest')
                ttv = item.get('PE', {}).get('totalTradedVolume')               
                # print(f"Strike Price: {strike}, Expiry: {expiry}, Call LTP: {ltp}")
                return [eq, uv, stk, exp, typ, ltp, iv, oi, coi, ttv]
            else:
                print("Invalid option type. Use 'CE' or 'PE'.")
                return None
        # print(item['strikePrice'])
        # print(item.get('strikePrice'))

    print("Option data not found for the given strike and expiry.")
    return None

# print(eq_opt_ltp("SBIN", 700, '26-Jun-2025', 'PE'))

def indc_opt_ltp(indc, stk, exp, typ):
    option_data = nse_indc_option(indc)['records']['data']

    for item in option_data:
        strike = item.get('strikePrice')
        expiry = item.get('expiryDate')
        
        # print(stk,strike)
        # print(exp,expiry)
        if stk == strike and expiry == exp:
            if typ == 'CE':
                uv = item.get('CE', {}).get('underlyingValue')
                ltp = item.get('CE', {}).get('lastPrice')
                iv = item.get('CE', {}).get('impliedVolatility')
                oi = item.get('CE', {}).get('openInterest')
                coi = item.get('CE', {}).get('changeinOpenInterest')
                ttv = item.get('CE', {}).get('totalTradedVolume')               
                # print(f"Strike Price: {strike}, Expiry: {expiry}, Call LTP: {ltp}")
                return [indc, uv, stk, exp, typ, ltp, iv, oi, coi, ttv]
            elif typ == 'PE':
                uv = item.get('PE', {}).get('underlyingValue')
                ltp = item.get('PE', {}).get('lastPrice')
                iv = item.get('PE', {}).get('impliedVolatility')
                oi = item.get('PE', {}).get('openInterest')
                coi = item.get('PE', {}).get('changeinOpenInterest')
                ttv = item.get('PE', {}).get('totalTradedVolume')               
                # print(f"Strike Price: {strike}, Expiry: {expiry}, Call LTP: {ltp}")
                return [indc, uv, stk, exp, typ, ltp, iv, oi, coi, ttv]
            else:
                print("Invalid option type. Use 'CE' or 'PE'.")
                return None
        # print(item['strikePrice'])
        # print(item.get('strikePrice'))

    print("Option data not found for the given strike and expiry.")
    return None

#print(indc_opt_ltp('NIFTY',25000,'29-May-2025','CE'))