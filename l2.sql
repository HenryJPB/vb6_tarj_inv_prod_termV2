SELECT EN01CANC.Codigo, EN01CANC.Referencia, EN01CANC.CodigoBanco, 
       EN01CANC.CodigoClienteProveedor, EN01CANC.ReconoceComoAnticipo,
  FROM EN01CANC
 WHERE EN01CANC.CodigoClienteProveedor = '11261122'
   AND EN01CANC.ReconoceComoAnticipo = 1
/