import requests

cookies = {
    'ubid-main': '134-2916247-3956466',
    'sid': '"RKuvscDzYqo2pj6EC6xZSg==|rFEMF0Ui1EtgdZNsSaK0if3w96U0CrXu+BuN5pNtC8w="',
    '__Host-mselc': 'H4sIAAAAAAAA/6tWSs5MUbJSSsytyjPUS0xOzi/NK9HLT85M0XM0DA0KDncOcgly8bWMUNJRykVSmZtalJyRCFKq52jkZ+pp6GTi6Gvq5OkGUpeNrLAApCQkLMDF29M7wsDFNUipFgAKAR4EdQAAAA==',
    'lc-main': 'en_US',
    'i18n-prefs': 'USD',
    'session-id': '131-7427739-0234048',
    'sp-cdn': '"L5Z9:CA"',
    'csm-hit': 'tb:BCSM1CGNC5V9WNVAWCFE+s-BCSM1CGNC5V9WNVAWCFE|1687961922928&t:1687961922928&adb:adblk_no',
    'session-id-time': '2318681929l',
    'x-main': '"dTPGR1EUeOl0rlwt6KYxgWgy@y79zqGILIqaEwhJZJFeO8DwqwiEUwoJHIt9cdD1"',
    'at-main': 'Atza|IwEBIOYwB3Vw6gSC-9mrn55LRLkEb0EM6phy53Pf-yBPU-gIp3NpvkPqV66XMcrULpL-D95n_V0K8y3utdhHU-Kjsycc-hGCVuiOZeV5ITbGjdXoCHzjAj5ykd8_Ydu1XQPf0zaZQoDpCuxljuuNhAKxAdkQxHU_9AC1PwXMKX0IFM1jmUqQhFI8HBd1Pfdrs9mV669wPW8zRwN2ETDIbBOqLQAS_gDLtXke7rpW5F3Bep_DQw',
    'sess-at-main': '"ZDpcJt3EExmE+bz3zjzLyOFjPxJEGGF6IL0BaaOuPHU="',
    'sst-main': 'Sst1|PQFl5t7iExTeRMvRjVqBF5t-CSMhh_H1suyQH_6sq0wbPbSp7fL5y2iGCApDjRauZFPSdgzV6epj77x_BnUAfaxnLK6k31ZhTLJbYVTgj-gexZtRFpuqc8XJ_4erx_DODYlWYnyXg_L1G5tVXNu-KXIrPz8m9DstbqGn6soXNdt3A1LxOhi1rpJdYBxgDVRoULtQQQBCw2akA6CcIXo5vLkVO_C1V1cTaxat6lw2rpxLVHIb6oshSnFoIlxiub9IkVkBbQlH1VbGR6RmGidkf9-B42sIQL9vtWI-IiCNVGzdNjE',
    'session-token': '2DTNi0ZBJPOXNfHdVnmJT++n8NGd/FByXZL6arK8MkFHSG9t+s89903k3kiRSIpQ5/GvFDQ3G48tw2a5+Af6SLk+Rqq1TG8qGP1M8IRe+R6gtrjlzBUafONShhI3qoV9QVnP7L2xaVfzaecAtjdEKm3B9qAQBEGvsiADocL/O/DjTpX+lTaGep15hDI5NYjGPpGdcDSlFktM7AiP2x7nAwieiFAybkqOD3z86y5Vr6r/0MCGx8puaUutfm9uta8CRpS8wzn2Nmk=',
}

headers = {
    'authority': 'sellercentral.amazon.com',
    'accept': '*/*',
    'accept-language': 'en-US,en;q=0.9',
    # 'cookie': 'ubid-main=134-2916247-3956466; sid="RKuvscDzYqo2pj6EC6xZSg==|rFEMF0Ui1EtgdZNsSaK0if3w96U0CrXu+BuN5pNtC8w="; __Host-mselc=H4sIAAAAAAAA/6tWSs5MUbJSSsytyjPUS0xOzi/NK9HLT85M0XM0DA0KDncOcgly8bWMUNJRykVSmZtalJyRCFKq52jkZ+pp6GTi6Gvq5OkGUpeNrLAApCQkLMDF29M7wsDFNUipFgAKAR4EdQAAAA==; lc-main=en_US; i18n-prefs=USD; session-id=131-7427739-0234048; sp-cdn="L5Z9:CA"; csm-hit=tb:BCSM1CGNC5V9WNVAWCFE+s-BCSM1CGNC5V9WNVAWCFE|1687961922928&t:1687961922928&adb:adblk_no; session-id-time=2318681929l; x-main="dTPGR1EUeOl0rlwt6KYxgWgy@y79zqGILIqaEwhJZJFeO8DwqwiEUwoJHIt9cdD1"; at-main=Atza|IwEBIOYwB3Vw6gSC-9mrn55LRLkEb0EM6phy53Pf-yBPU-gIp3NpvkPqV66XMcrULpL-D95n_V0K8y3utdhHU-Kjsycc-hGCVuiOZeV5ITbGjdXoCHzjAj5ykd8_Ydu1XQPf0zaZQoDpCuxljuuNhAKxAdkQxHU_9AC1PwXMKX0IFM1jmUqQhFI8HBd1Pfdrs9mV669wPW8zRwN2ETDIbBOqLQAS_gDLtXke7rpW5F3Bep_DQw; sess-at-main="ZDpcJt3EExmE+bz3zjzLyOFjPxJEGGF6IL0BaaOuPHU="; sst-main=Sst1|PQFl5t7iExTeRMvRjVqBF5t-CSMhh_H1suyQH_6sq0wbPbSp7fL5y2iGCApDjRauZFPSdgzV6epj77x_BnUAfaxnLK6k31ZhTLJbYVTgj-gexZtRFpuqc8XJ_4erx_DODYlWYnyXg_L1G5tVXNu-KXIrPz8m9DstbqGn6soXNdt3A1LxOhi1rpJdYBxgDVRoULtQQQBCw2akA6CcIXo5vLkVO_C1V1cTaxat6lw2rpxLVHIb6oshSnFoIlxiub9IkVkBbQlH1VbGR6RmGidkf9-B42sIQL9vtWI-IiCNVGzdNjE; session-token=2DTNi0ZBJPOXNfHdVnmJT++n8NGd/FByXZL6arK8MkFHSG9t+s89903k3kiRSIpQ5/GvFDQ3G48tw2a5+Af6SLk+Rqq1TG8qGP1M8IRe+R6gtrjlzBUafONShhI3qoV9QVnP7L2xaVfzaecAtjdEKm3B9qAQBEGvsiADocL/O/DjTpX+lTaGep15hDI5NYjGPpGdcDSlFktM7AiP2x7nAwieiFAybkqOD3z86y5Vr6r/0MCGx8puaUutfm9uta8CRpS8wzn2Nmk=',
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
    'countryCode': 'CA',
    'asin': 'B007BUS9GE',
    'fnsku': '',
    'searchType': 'GENERAL',
    'locale': 'en-US',
}

response = requests.get(
    'https://sellercentral.amazon.com/revenuecalculator/getadditionalpronductinfo',
    params=params,
    cookies=cookies,
    headers=headers,
)

print(response.json())