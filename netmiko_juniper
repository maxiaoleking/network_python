from netmiko import ConnectHandler

juniper = {
    'device_type': 'juniper',
    'host': '172.16.0.40',
    'username': 'root',
    'password': 'aruba123',
    'port' : 22,
}

net_connect = ConnectHandler(**juniper)

print('success connect to ' + juniper['ip'])

output = net_connect.send_command('show configuration')


print(output)


net_connect.disconnect()
