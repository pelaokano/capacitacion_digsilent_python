def resultados(elemento, variables):
    aux = [elemento.GetAttribute(var) if elemento.HasAttribute(var) else 0 for var in variables ]
    return aux 