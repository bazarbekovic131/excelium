# Here i will add lists of coordinators for different objects







def check_company_object_pair(company, object):
    # Perform some checks on the company and object pair
    # and return a list of integers
    
    # Example implementation:
    company_object_pairs = {
        ("Company A", "Object 1"): [1, 2, 3],
        ("Company B", "Object 2"): [4, 5, 6]
    }
    
    return company_object_pairs.get((company, object), [])