## 1. Usage
```sh
# 1. Environment
python -m venv .venv
source .venv/bin/activate
pip3 install -r requirements.txt


cd proj
./manage.py runserver
```

## 2. Test
1. Login
    POST: http://127.0.0.1:8000/api/token/
    ```json
    {
    "username": "admin",
    "password": "Password123"
    }
    ```

    Response: e.g.
    ```json
    {
        "refresh": "eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJ0b2tlbl90eXBlIjoicmVmcmVzaCIsImV4cCI6MTY2MDI3NjUwMSwiaWF0IjoxNjYwMTkwMTAxLCJqdGkiOiIzZjY0ZjZiMWZmMmM0NTBiYWZmYmM1MGQyODA1YmU1ZSIsInVzZXJfaWQiOjF9.sxlefGbwpzOxJJitb2-7925uX9GbytbEU0-K7Tjc0SU",
        "access": "eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJ0b2tlbl90eXBlIjoiYWNjZXNzIiwiZXhwIjoxNjYwMTkwNDAxLCJpYXQiOjE2NjAxOTAxMDEsImp0aSI6ImQyNmRjNDZlODFkOTRmY2FhM2YwOTNhZjAzYWU0ZDQ0IiwidXNlcl9pZCI6MX0.UqTLo1pgkR1e699OPl_X0mvHjS0BVz1liUJoTEcgZLM"
    }
    ```
2. Get Customers list
    GET: http://127.0.0.1:8000/api/customers/

    Header: 
    ```    
    Authorization: "Bearer <access token>"
    ```
3. Create a Customer
    POST: http://127.0.0.1:8000/api/customers/
    
    Header: 
    ```
    Authorization: "Bearer <access token>"
    ```
4. GET a Customer
    GET: http://127.0.0.1:8000/api/customers/<customer id>/

    Header: 
    ```
    Authorization: "Bearer <access token>"
    ```
5. Edit a Customer
    PATCH: http://127.0.0.1:8000/api/customers/<customer id>/
    
    Body:
    ```json
    {
        "first_name": "walk dog1222",
        "last_name": "Second last name1222",
        "address": "second address1222"
    }
    ```
    Header: 
    ```
    Authorization: "Bearer <access token>"
    ```
6. Delete a Customer
    DELETE: http://127.0.0.1:8000/api/customers/<customer id>/

    Header: 
    ```
    Authorization: "Bearer <access token>"
    ```
