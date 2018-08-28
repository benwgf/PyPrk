class Base(object):
   
    def __init__(self):
        self.__name=type(self).__name__
        pass

    def sayHello(self):
        print("hello from "+self.__name)

    @property
    def Name(self):
        return self.__name
    @Name.setter
    def Name(self,value):
        self.__name=value
    

    @staticmethod
    def sfunc():
        print("invoke statics method")
    @classmethod
    def cfunc(cls):
        print("invoke class method")
    

class B(Base):
    pass

b=B()
print(b.Name)
b.Name="hh"
print(b.Name)
b.sayHello()
b.sfunc()
B.sfunc()
b.cfunc()
B.cfunc()
