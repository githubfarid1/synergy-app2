import requests

cookies = {
    'session-id': '143-3445911-9496428',
    'i18n-prefs': 'USD',
    'ubid-main': '135-1698517-0362129',
    's_pers': '%20s_fid%3D71A5BA2BB54387EC-23184C6241077C36%7C1843500634686%3B%20s_dl%3D1%7C1685649634687%3B%20s_ev15%3D%255B%255B%2527NSGoogle%2527%252C%25271685647834690%2527%255D%255D%7C1843500634690%3B',
    'sid': '"dX8/EHJKQIXIjQiHhXasMw==|WqlaeXcqyUHpafXiPOE+HvuAXb9VtooobpWbpBGPdRw="',
    '__Host-mselc': 'H4sIAAAAAAAA/6tWSs5MUbJSSsytyjPUS0xOzi/NK9HLT85M0XM0DA0KDncOcgly8bWMUNJRykVSmZtalJyRCFKq52jkZ+pp6GTi6Gvq5OkGUpeNrLAApCQkLMDF29M7wsDFNUipFgAKAR4EdQAAAA==',
    'csm-hit': 'tb:PX1964Z7C79TD2WTG6SF+s-PX1964Z7C79TD2WTG6SF|1686870448919&t:1686870448919&adb:adblk_no',
    'x-main': '"gtbOEYZN@OUowIfunUGsjiVHUzltTTkVWqILyfswxxe2ajrdhVUiZ7gHsPyMs6rM"',
    'at-main': 'Atza|IwEBIGhTt9SC20bEDL_jPZ-uWJ8eOPMFcghh2361jPFWp6na_i0MegZBxkkd41Bpp8WrKxNWtcRZfmF2S9dt8P-EoGpqCqnjV2mDjzd3SCjaaYhSSIprXyykVKpB9l0ZV4ygu5t4XwMc69tPZcbVl9wauMC7jGNhwoYfUtyVifeFjGgqicfSUGwaLOI40Cr_yHMzasSZ-QypwZ6Dd_mRAkt8oHTQqTZePHDMM0-AthhN2otI6Q',
    'sess-at-main': '"sWzNJufEovEvfA5e7MN4evAhPItNf7P30NyYRq6arZQ="',
    'sst-main': 'Sst1|PQEgEQOY7GkT5bLupUDyHd_JCZ0-Y3uGRb_jlkpn-M3PdlFyQ4xQD2LZM5TN_NBF0qOg9tUDb0-BdOnHkDhUfTfkGABaziFvpQ2RoOKsFz7nSCLlz32DlqBFns5La5L_GC5pc1lfEB1v_AHnIc45fzMxGC-o2UqGSxaTYiGDrVNYfcBecWCei0CeafmHa1P9LwRLN__lTQ2KzEKIVjlS4mBIZdJF_mk7wR7qB7jWD2-qigz4ATyuzzEtNcvKzMXolMMvtp0xbODE8_KHtFlJNlMB2JCiUQ1vnHKCm-xQnUMOmug',
    'session-id-time': '2082787201l',
    'session-token': 'laGVoAf8uOh/IUcntb9QA/S4yyT0pqgo1N+C/TwehlQMpZetIa0yFWCy3fVMdnjEYy/phx3QJtCL/9TUdbb3CMwNzLK5D1RrqrUuS1CSTdRJ7AooIEG4MLvF0cNHroH6p58LMBV2oUAxUNmNe33XnrP+lMs6AHHvp+vGtp1pnFYTad67LTPPY0hZcbfH/iudg96f/VPd0oomop/govpyPpng2VY8wctn3HlYGPL7m+pQsavM2kA+4xFa2dgyWNU3k0gZXBW99gA=',
}

headers = {
    'authority': 'sellercentral.amazon.com',
    'accept': '*/*',
    'accept-language': 'en-US,en;q=0.9',
    # 'cookie': 'session-id=143-3445911-9496428; i18n-prefs=USD; ubid-main=135-1698517-0362129; s_pers=%20s_fid%3D71A5BA2BB54387EC-23184C6241077C36%7C1843500634686%3B%20s_dl%3D1%7C1685649634687%3B%20s_ev15%3D%255B%255B%2527NSGoogle%2527%252C%25271685647834690%2527%255D%255D%7C1843500634690%3B; sid="dX8/EHJKQIXIjQiHhXasMw==|WqlaeXcqyUHpafXiPOE+HvuAXb9VtooobpWbpBGPdRw="; __Host-mselc=H4sIAAAAAAAA/6tWSs5MUbJSSsytyjPUS0xOzi/NK9HLT85M0XM0DA0KDncOcgly8bWMUNJRykVSmZtalJyRCFKq52jkZ+pp6GTi6Gvq5OkGUpeNrLAApCQkLMDF29M7wsDFNUipFgAKAR4EdQAAAA==; csm-hit=tb:PX1964Z7C79TD2WTG6SF+s-PX1964Z7C79TD2WTG6SF|1686870448919&t:1686870448919&adb:adblk_no; x-main="gtbOEYZN@OUowIfunUGsjiVHUzltTTkVWqILyfswxxe2ajrdhVUiZ7gHsPyMs6rM"; at-main=Atza|IwEBIGhTt9SC20bEDL_jPZ-uWJ8eOPMFcghh2361jPFWp6na_i0MegZBxkkd41Bpp8WrKxNWtcRZfmF2S9dt8P-EoGpqCqnjV2mDjzd3SCjaaYhSSIprXyykVKpB9l0ZV4ygu5t4XwMc69tPZcbVl9wauMC7jGNhwoYfUtyVifeFjGgqicfSUGwaLOI40Cr_yHMzasSZ-QypwZ6Dd_mRAkt8oHTQqTZePHDMM0-AthhN2otI6Q; sess-at-main="sWzNJufEovEvfA5e7MN4evAhPItNf7P30NyYRq6arZQ="; sst-main=Sst1|PQEgEQOY7GkT5bLupUDyHd_JCZ0-Y3uGRb_jlkpn-M3PdlFyQ4xQD2LZM5TN_NBF0qOg9tUDb0-BdOnHkDhUfTfkGABaziFvpQ2RoOKsFz7nSCLlz32DlqBFns5La5L_GC5pc1lfEB1v_AHnIc45fzMxGC-o2UqGSxaTYiGDrVNYfcBecWCei0CeafmHa1P9LwRLN__lTQ2KzEKIVjlS4mBIZdJF_mk7wR7qB7jWD2-qigz4ATyuzzEtNcvKzMXolMMvtp0xbODE8_KHtFlJNlMB2JCiUQ1vnHKCm-xQnUMOmug; session-id-time=2082787201l; session-token=laGVoAf8uOh/IUcntb9QA/S4yyT0pqgo1N+C/TwehlQMpZetIa0yFWCy3fVMdnjEYy/phx3QJtCL/9TUdbb3CMwNzLK5D1RrqrUuS1CSTdRJ7AooIEG4MLvF0cNHroH6p58LMBV2oUAxUNmNe33XnrP+lMs6AHHvp+vGtp1pnFYTad67LTPPY0hZcbfH/iudg96f/VPd0oomop/govpyPpng2VY8wctn3HlYGPL7m+pQsavM2kA+4xFa2dgyWNU3k0gZXBW99gA=',
    'referer': 'https://sellercentral.amazon.com/revcal?ref=RC1&',
    'sec-ch-ua': '"Not.A/Brand";v="8", "Chromium";v="114", "Google Chrome";v="114"',
    'sec-ch-ua-mobile': '?0',
    'sec-ch-ua-platform': '"Windows"',
    'sec-fetch-dest': 'empty',
    'sec-fetch-mode': 'cors',
    'sec-fetch-site': 'same-origin',
    'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36',
}

params = {
    'searchKey': 'B0187HZG2E',
    'countryCode': 'CA',
    'locale': 'en-US',
}

response = requests.get(
    'https://sellercentral.amazon.com/revenuecalculator/productmatch',
    params=params,
    cookies=cookies,
    headers=headers,
)

print(response.json())