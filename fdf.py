import random

data = []
for i in range(5):
    data.append({
        "{{Product Count}}": i+1,
        "{{Product}}": "Product " + str(i+1),
        "{{Brand}}": "Brand " + str(i+1),
        "{{Specification}}": "Specification " + str(i+1),
        "{{Purchase Qty.}}": random.randint(1, 10),
        "{{Unit}}": "Unit " + str(i+1),
        "{{Unit Price}}": random.uniform(1, 100),
        "{{Net Value}}": random.uniform(1, 1000),
        "{{Tax Rate%}}": random.uniform(1, 10),
        "{{GST}}": random.uniform(1, 100),
        "{{Total Amount}}": random.uniform(1, 10000)
    })

print(data)
