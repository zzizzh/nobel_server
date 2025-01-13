from pylogix import PLC

# PLC의 IP 주소 설정
plc_ip = "192.168.1.10"  # PLC의 실제 IP 주소로 대체하세요.

# EtherNet/IP 통신 설정 및 데이터 읽기
with PLC() as comm:
    comm.IPAddress = plc_ip

    # 5000번 주소의 flag 값을 읽기
    flag_tag = "Flag"  # PLC에서 5000번 주소를 의미하는 태그 이름
    flag_value = comm.Read(flag_tag).Value

    print(f"Flag 값: {flag_value}")

    # flag 값이 1 또는 2일 때만 데이터 읽기
    if flag_value in [1, 2]:
        # 5002번부터 17개의 DoubleWord 읽기
        data_tags = [f"Data{i}" for i in range(5002, 5002 + 17)]
        data_values = [comm.Read(tag).Value for tag in data_tags]

        print("읽은 데이터:")
        for i, value in enumerate(data_values, start=5002):
            print(f"Address {i}: {value}")
    else:
        print("Flag 값이 0이므로 데이터를 읽지 않습니다.")