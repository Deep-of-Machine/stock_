a = ['A373220', 'A000660', 'A207940', 'A005935', 'A035420', 'A005380', 'A006400', 'A000270', 'A055550', 'A034730', 'A066570', 'A323410', 'A015760', 'A247540', 'A011200', 'A329180', 'A051900', 'A259960', 'A032830', 'A034020', 'A017670', 'A086790', 'A009150', 'A003670', 'A010950', 'A302440', 'A003490', 'A066970', 'A316140', 'A009830', 'A377300', 'A352820', 'A086280', 'A011170', 'A361610', 'A047810', 'A035250', 'A034220', 'A032640', 'A088980', 'A011790', 'A000720', 'A021240', 'A267250', 'A004020', 'A000100', 'A293490', 'A006800', 'A000060', 'A271560', 'A011780', 'A028050', 'A078930', 'A004990', 'A128940', 'A029780', 'A371460', 'A071050', 'A307950', 'A263750', 'A012450', 'A282330', 'A068760', 'A016360', 'A028670', 'A008560']

b = ['A000660', 'A207940', 'A005935', 'A051910', 'A035420', 'A005380', 'A035720', 'A028260', 'A005490', 'A096770', 'A034730', 'A066570', 'A247540', 'A329180', 'A030200', 'A000810', 'A003490', 'A010130', 'A066970', 'A377300', 'A024110', 'A352820', 'A086280', 'A383220', 'A047810', 'A018880', 'A021240', 'A004020', 'A293490', 'A006800', 'A011780', 'A078930', 'A128940', 'A371460', 'A263750', 'A138040', 'A012450', 'A282330', 'A241560', 'A008560']
def intersect(a, b):
    return list(set(a) & set(b))

print(intersect(a,b))