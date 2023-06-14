from src.interface import ParseAPIDoc



if __name__ == '__main__':
    p = ParseAPIDoc()
    project = p.run('swagger.html')
