from .ADODB import Connection
from comtypes import client
from collections import deque

DRIVER = "Oracle en OraClient11g_home1"
UID = "cobros_sox"
PWD = "cobros_sox"
DBQ = "ADMCATP"
SERVERS = {
        1: "ADMCATP",
        2: "ADMCATP",
        3: "CUST01P",
        4: "CUST02P",
        5: "CUST03P",
        6: "CUST04P",
        7: "CUST05P",
        8: "CUST06P",
        9: "CUST90P",
        10: "CUSTHIP"
        }
CONNECTION_STRING = "Driver={};uid={};pwd={};dbq={}".format(
        DRIVER,
        UID,
        PWD,
        DBQ
        )
LETRAS = [letra for letra in
        "qwertyuiopasdfghjklzxcvbnm"
        ]
NUMEROS = [letra for letra in
        "1234567890"
        ]

class Rowset(object):
    def __init__(self, data):
        self._data = data
        self._count = None
        self._index = int()
        self._columns = list()
        self._iter = int()
    
    def __getattr__(self, attr):
        try:
            dato = self.data.Fields.Item(
                    self.columns.index(attr.lower())
                    ).Value
        except:
            dato = None
        return dato

    def __getitem__(self, index):
        try:
            self.data.MoveFirst()
        except:
            return None
        self._index = index
        for x in range(index):
            try:
                self.data.MoveNext()
            except:
                if x == 0: 
                    self._count = 0
                break
        return self

    '''def __del__(self):
        try:
            self.data.Release()
        except:
            pass'''

    def __repr__(self):
        final = str()
        final += "\n{}\n".format("-"*40)
        for column in self.columns:
            final += "{}: {}\n".format(column, self.__getattr__(column))
        final += "{}\n".format("-"*40)
        return final

    def __iter__(self):
        self._iter = 0
        return self

    def __next__(self):
        try:
            if self._iter < self.count:
                data = self[self._iter]
                self._iter += 1
                return data
            else:
                raise Exception
        except:
            try:
                self.data.MoveFirst()
            except:
                pass
            raise StopIteration

    @property
    def columns(self):
        if self._columns == list():
            for x in range(self.data.Fields.Count):
                self._columns.append(
                        self.data.Fields.Item(x).Name.lower()
                        )
        return self._columns

    @property
    def count(self):
        if self._count == None:
            try:
                self.data.MoveFirst()
            except:
                pass
            self._count = int()
            while True:
                if self.data.EOF:
                    if self._count > 0:
                        self.data.MoveFirst()
                        self[self._index]
                    break
                else:
                    self._count += 1
                    self.data.MoveNext()
        return self._count
        
    @property
    def data(self):
        return self._data[1]

        
class App(object):   
    def __init__(self):
        self._connection = client.CreateObject(Connection)
        self.connection.ConnectionString = CONNECTION_STRING
        self.connection.open()

    def __del__(self):
        self.connection.close()

    @property
    def connection(self):
        return self._connection

    def execute(self, sql, bind_list=tuple()):
        sql = App.parse_sql(sql, bind_list)
        data = self.connection.execute(sql)
        return Rowset(data)

    def get_ssn(self, ssn):
        final = dict()
        letra = str()
        if ssn[0].lower() in LETRAS and ssn[0].lower() != "x":
            letra = "C"
        elif ssn[0].lower() != "x":
            letra = "N"
        elif ssn[0] in NUMEROS and ssn[-1].lower() in LETRAS:
            letra = "N"
        else:
            letra = "P"
        datos = self.execute(
                "select * from external_id_acct_map where external_id like '?%' order by external_id asc",
                ("{}{}".format(letra, ssn), )
                )
        for dato in datos:
            external_id = dato.external_id
            datos_finales = self.get_balances(external_id)
            final.update(datos_finales)
        return final #Devolver clase CUSTOMER

    def get_external_id(self, external_id):
        final = dict()
        datos = self.execute(
                "select * from external_id_acct_map where external_id = ? order by external_id asc",
                (external_id, )
                )
        for dato in datos:
            datos_finales = self.execute(
                    "select * from cmf@{} where account_no=? order by account_no asc".format(
                            SERVERS[dato.server_id].lower()
                            ),
                    (dato.account_no, )
                    )
            subscr_no = self.execute(
                    "select subscr_no, subscr_no_resets from service@{} where parent_account_no = ? order by parent_account_no asc".format(
                            SERVERS[dato.server_id].lower()
                            ),
                    (dato.account_no, )
                    )
            telefonos = self.execute(
                    "select * from customer_id_equip_map@{} where subscr_no=? and subscr_no_resets=? and inactive_date is null and external_id_type=1 order by subscr_no asc".format(
                            SERVERS[dato.server_id].lower()
                            ),
                    (subscr_no.subscr_no, subscr_no.subscr_no_resets)
                    )
            servicios = self.execute(
                    "select * from cmf_package_component@{0} inner join component_definition_values@{0} on cmf_package_component.component_id = component_definition_values.component_id where parent_account_no = ?".format(
                            SERVERS[dato.server_id].lower()
                            ),
                    (dato.account_no,)
                    )
            if datos_finales.count > 0:
                final[external_id] = {
                        "lookout": datos, 
                        "cmf": datos_finales,
                        "phones": telefonos,
                        "services": servicios
                        }
        return final #Devolver clase CUSTOMER

    def get_balances(self, external_id):
        cliente = self.get_external_id(external_id)
        account_no = cliente[external_id]["cmf"].account_no
        server = SERVERS[cliente[external_id]["lookout"].server_id]
        balances = self.execute(
                "select * from cmf_balance@{} where account_no=? order by account_no asc".format(
                        server.lower()
                        ),
                (account_no, )
                )
        cliente[external_id]["balances"] = Balances(balances)
        return cliente

    @staticmethod
    def parse_sql(sql, bind_list=tuple()):
        sql = sql.replace("\n", " ")
        sql = sql.replace(";", "")
        sql_pieces = sql.split("?")
        sql_final = str()
        if len(bind_list) > 0:
            for index, piece in enumerate(sql_pieces):
                if index < len(sql_pieces)-1:
                    if "like" in piece:
                        sql_final += "{}{}".format(piece, str(bind_list[index]).replace("'", "''").replace("\"", "\"\"")) #Hacer que verifique si tiene o no comillas
                    else:
                        sql_final += "{}'{}'".format(piece, str(bind_list[index]).replace("'", "''").replace("\"", "\"\""))
                else:
                    sql_final += piece
        if sql_final==str():
            sql_final = sql
        return sql_final

class Customer(object):
    def __init__(self, app=App(), cache=10):
        self._app = App()
        self._datos = dict()
        self._ssn = str()
        self._cache = Cache(cache)

    def __getitem__(self, item):
        self.search(item)
        if isinstance(item, int):
            return self._datos[self.external_ids[item]]
        elif isinstance(item, str) and item in self.external_ids:
            return self._datos[item]
        elif isinstance(item, str) and item==self.ssn:
            return self
        else:
            raise Exception

    def search(self, ssn):
        self._ssn = ssn
        self._datos = self._cache.get(ssn)
        if not self._datos:
            if len(ssn) > 10:
                ssn = ssn[1:-4]
            self._datos = self.app.get_ssn(ssn)
            for dato in self._datos:
                self._datos[dato] = ExternalId(self._datos[dato])
            self._cache.insert(self._ssn, self._datos)

    def load_external_id(self, external_id): #Pa borrar
        return ExternalId(self.app.get_external_id(external_id))

    @property
    def ssn(self):
        return self._ssn
    
    @property
    def external_ids(self):
        return [dato for dato in self._datos]

    @property
    def app(self):
        return self._app

    @property
    def due(self):
        return sum([self.__getitem__(item).due for item in self.external_ids])

    @property
    def phones(self):
        phones = list()
        for item in self.external_ids:
            phones.extend(self.__getitem__(item).phones)
        return phones

class ExternalId(object):
    def __init__(self, datos):
        self._datos = datos
        self._due = None
        self._phones = None

    @property
    def cmf(self):
        return self._datos["cmf"]

    @property
    def balances(self):
        return self._datos["balances"]

    @property
    def lookout(self):
        return self._datos["lookout"]

    @property
    def external_id(self):
        return self.lookout.external_id

    @property
    def name(self):        
        return "{} {}".format(self.cmf.bill_fname, self.cmf.bill.lname)

    @property
    def phone1(self):
        return self.cmf.phone1

    @property
    def phone2(self):
        return self.cmf.phone2

    @property
    def due(self):
        if self._due == None:
            self._due = sum([bill.balance_due for bill in self.balances])/100
        return self._due

    @property
    def phones(self):
        if self._phones == None:
            self._phones = list()
            for item in self._datos["phones"]:
                self._phones.append(item.external_id)
        return self._phones

    def bill(self, ref_no):
        if ref_no in self.balances:
            return {
                    "total_due": self.balances[ref_no].total_due/100,
                    "total_adj": self.balances[ref_no].total_adj/100,
                    "total_paid": self.balances[ref_no].total_paid/100,
                    "balance_due": self.balances[ref_no].balance_due/100
                    }

class Balances(object):
    def __init__(self, balances):
        self._balances = balances
        self._index = None

    def __iter__(self):
        self._iter = 0
        return self

    def __next__(self):
        try:
            if self._iter < self.count:
                data = self[self._iter]
                self._iter += 1
                return data
            else:
                raise Exception
        except:
            self._iter = 0
            raise StopIteration

    @property
    def index(self):
        if self._index == None:
            return [bill.bill_ref_no for bill in self._balances]
        else:
            return self._index
    

    def __getitem__(self, item):
        if isinstance(item, int):
            return self._balances[item]
        elif isinstance(item, str) and item in self.index:
            return self._balances[self.index.index(item)]

    @property
    def count(self):
        return self._balances.count

class Cache(object):
    def __init__(self, maximum=10):
        self._maximum = maximum
        self._deque = deque([], maximum)
        self._dict = dict()

    @property
    def list(self):
        return list(self._deque)

    @property
    def maximum(self):
        return self._maximum

    def get(self, key):
        if key in self._deque:
            self._deque.remove(key)
            self._deque.append(key)
            return self._dict[key]
        else:
            return False

    def insert(self, key, item):
        if item in self._deque:
            self._deque.remove(key)
        self._deque.append(key)
        self._dict[key] = item
        for key in self._deque:
            if key not in self._dict:
                del self._dict[key]

    @property
    def next(self):
        self._deque.rotate(-1)
        return self._deque[0]
        
    @property
    def previous(self):
        self._deque.rotate(1)
        return self._deque[0]
