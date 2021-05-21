a = {'name': 'Zhangsan', 'age': 18}
b = ['Wangwu', 22]


class Person:
    def __init__(self, **kwargs):
        self.name = kwargs.get('name')
        self.age = kwargs.get('age')


if __name__ == '__main__':
    # p = Person()
    # for k, v in a.items():
    #     setattr(p, k, v)
    p = Person(**a)
    print(p.name)
    print(p.age)

