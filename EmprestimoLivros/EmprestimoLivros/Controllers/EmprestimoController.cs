using ClosedXML.Excel;
using EmprestimoLivros.Data;
using EmprestimoLivros.Models;
using Microsoft.AspNetCore.Mvc;
using System.Data;

namespace EmprestimoLivros.Controllers
{
    public class EmprestimoController : Controller
    {
        //puxando a classe que ja esta conectada com o banco
        readonly private ApplicationDbContext _db; // o underline e pq e padrao de nomenclatura para propiedades privadas

        //construtor
        public EmprestimoController(ApplicationDbContext db)
        {
            _db = db;
        }


        public IActionResult Index()
        {
            IEnumerable<EmprestimoModel> emprestimos = _db.Emprestimos; // pegando a tabela do banco
            return View(emprestimos);
        }

        [HttpGet] // quando tem isso e GET
        public IActionResult Cadastrar()
        {
            return View();
        }

        [HttpGet]// quando tem isso e GET
        public IActionResult Editar(int? id)
        {
            if (id == null || id == 0)
            {
                return NotFound();
            }

            EmprestimoModel emprestimo = _db.Emprestimos.FirstOrDefault(x => x.Id == id); // esse codigo e igual a um select com where (select * from table where id = id enviado

            if (emprestimo == null)
            {
                return NotFound();
            }


            return View(emprestimo);
        }

        [HttpGet]

        public IActionResult Excluir(int? id)
        {
            if (id == null || id == 0)
            {
                return NotFound();
            }

            EmprestimoModel emprestimo = _db.Emprestimos.FirstOrDefault(x => x.Id == id); // esse codigo e igual a um select com where (select * from table where id = id enviado

            if (emprestimo == null)
            {
                return NotFound();
            }


            return View(emprestimo);
        }

        [HttpGet]
        public IActionResult Exportar()
        {
            var dados = GetDados();

            using (XLWorkbook workbook = new XLWorkbook()) // usansado using pois queremos criar essa operacao utilizar ela e desaparecer com ela ,se nao usar using ela fica aberta
            {
                workbook.AddWorksheet(dados, "Dados Emprestimos");

                using (MemoryStream ms = new MemoryStream())
                {
                    workbook.SaveAs(ms);
                    return File(ms.ToArray(), "application/vnd.openxmlformats-officedocument.spredsheetml.sheet", "Empestimos.xls");
                }
            }
        }

        private DataTable GetDados() //esta privado pois apenas o metodo a cima vai utilizar esse de baixo
        {
            DataTable dataTable = new DataTable();

            dataTable.TableName = "Dados Emprestimos";
            dataTable.Columns.Add("Id", typeof(int));
            dataTable.Columns.Add("Recebedor", typeof(string));
            dataTable.Columns.Add("Fornecedor", typeof(string));
            dataTable.Columns.Add("Livro", typeof(string));
            dataTable.Columns.Add("Data Emprestimo", typeof(DateTime));

            var dados = _db.Emprestimos.ToList();
            if (dados.Count > 0)
            {
                dados.ForEach(emprestimos => dataTable.Rows.Add(emprestimos.Id, emprestimos.Recebedor, emprestimos.Fornecedor, emprestimos.LivroEmprestado, emprestimos.DataUltimaAtualizacao));
            }


            return dataTable;
        }


        [HttpPost] // quando tem isso e POST
        public IActionResult Cadastrar(EmprestimoModel emprestimo)
        {
            if (ModelState.IsValid)
            {
                _db.Emprestimos.Add(emprestimo);
                _db.SaveChanges();

                TempData["MessagemSucesso"] = "Cadastro Realizado com Sucesso!!";

                return RedirectToAction("Index"); // depois que tudo der certo manda o usuario para a view INDEX
            }
            TempData["MessagemErro"] = "Algo deu errado, Nenhuma alteracao Foi realizada!!!";
            return View();
        }

        [HttpPost]
        public IActionResult Editar(EmprestimoModel emprestimo)
        {
            if (ModelState.IsValid)
            {
                _db.Emprestimos.Update(emprestimo);
                _db.SaveChanges();

                TempData["MessagemSucesso"] = "Cadastro Editado com Sucesso!!";

                return RedirectToAction("Index");
            }
            TempData["MessagemErro"] = "Algo deu errado, Nenhuma alteracao Foi realizada!!!";
            return View(emprestimo);
        }

        [HttpPost]

        public IActionResult Excluir(EmprestimoModel emprestimo)
        {
            if (emprestimo == null)
            {
                TempData["MessagemErro"] = "Algo deu errado, Nenhuma alteracao Foi realizada!!!";
                return NotFound();
            }

            _db.Emprestimos.Remove(emprestimo);
            _db.SaveChanges();

            TempData["MessagemSucesso"] = "Cadastro Excluido com Sucesso!!";
            return RedirectToAction("Index");
        }


    }
}
