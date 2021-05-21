import ipaddress

a = ipaddress.ip_address('192.168.1.1')
b = a.packed
print(int.from_bytes(b, byteorder='big', signed=False))
c = ipaddress.ip_address(3232235777)
print(c)

net_work = ipaddress.ip_network('192.168.1.0/24')
print(net_work.hostmask)
print(list(net_work.hosts()))
print(c in net_work)