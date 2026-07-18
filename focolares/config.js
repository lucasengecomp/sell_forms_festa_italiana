// Manter PRECO em sincronia com Code.gs (o servidor recalcula os totais).
var WA_NUMBER = '5511987529809';
var PRECO = { adulto: 165, criancaPagante: 85 };

// Dados do recebedor PIX (limites do BR Code: nome 25 chars, cidade 15 chars)
var PIX = {
  chave:  '05056796000207',
  nome:   'MOVIMENTO DOS FOCOLARES',
  cidade: 'VARGEM GRANDE P'
};

var FORMAS_PAGAMENTO = {
  pix:    'PIX',
  cartao: 'Cartão de crédito',
  balcao: 'Balcão — Padaria Espiga Dourada',
  isento: 'Isento (sem valor a pagar)'
};
