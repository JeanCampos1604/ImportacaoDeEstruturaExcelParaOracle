using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections;

namespace TesteImportacaoExcel.Tabelas
{
    public class IdentificarPai
    {
        public List<Item> _itens;

        public IdentificarPai(List<Item> itens)
        {
            _itens = itens;
        }

        public void Processar(string identificador, Item itemPai)
        {
            // Itera sobre os itens na lista "_itens" onde a ordem é maior que a ordem do item pai
            foreach (Item item in _itens.Where(r => r.Ordem > itemPai.Ordem))
            {
                // Verifica se o item possui um código de pai definido e passa para a próxima iteração caso tenha
                if (!String.IsNullOrEmpty(item.CodPai))
                    continue;

                // Verifica se o código do produto do item contém o identificador fornecido
                // ou se o tipo do item é "MP" ou "MOT" ou se o código do produto começa com "999"
                if (item.CodProduto.Contains(identificador) ||
                    item.Tipo == "MP" ||
                    item.Tipo == "MOT" ||
                    item.CodProduto.StartsWith("999"))
                {
                    // Define o código do pai do item como o código do produto do item pai
                    item.CodPai = itemPai.CodProduto;
                }
                else
                {
                    // Obtém o item anterior com base na ordem do item pai
                    Item itemAnterior = _itens.First(r => r.Ordem == (item.Ordem));

                    // Verifica se o tipo do item anterior é "SEMI" e define o código do pai do item atual como o código do produto do item pai
                    if (itemAnterior.Tipo == "SEMI")
                        item.CodPai = itemPai.CodProduto;
                }

                // Se o tipo do item atual for "SEMI", chama recursivamente o método "Processar" com o código do produto do item pai e o item atual
                if (item.Tipo == "SEMI")
                    Processar(itemPai.CodProduto, item);
            }
        }

    }
}
