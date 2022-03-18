using Excel;
using Habitasorte.Business.Model;
using Habitasorte.Business.Model.Publicacao;
using Habitasorte.Business.Pdf;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data.SqlServerCe;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace Habitasorte.Business {

    public delegate void SorteioChangedEventHandler(Sorteio s);

    public class SorteioService {

        public event SorteioChangedEventHandler SorteioChanged;

        private Sorteio model;
        public Sorteio Model {
            get { return model; }
            set { model = value; SorteioChanged(model); }
        }

        public SorteioService() {
            Database.Initialize();
        }

        private void Execute(Action<Database> action) {
            using (SqlCeConnection connection = Database.CreateConnection()) {
                using (SqlCeTransaction tx = connection.BeginTransaction()) {
                    Database database = new Database(connection, tx);
                    try {
                        action(database);
                        tx.Commit();
                    } catch(Exception e) {
                        try { tx.Rollback(); } catch { }
                        throw;
                    }
                }
            }
        }

        private void AtualizarStatusSorteio(Database database, string status) {
            Model.StatusSorteio = status;
            database.AtualizarStatusSorteio(status);
        }

        /* Configuração */

        public void ExcluirBancoReiniciarAplicacao() {
            Database.ExcluirBanco();
            System.Windows.Application.Current.Shutdown();
        }

        /* Ações */

        public void CarregarSorteio() {
            Execute(d => {
                Model = d.CarregarSorteio();
            });
        }

        public void AtualizarSorteio() {
            Execute(d => {
                d.AtualizarSorteio(Model);
                AtualizarStatusSorteio(d, Status.IMPORTACAO);
            });
        }

        public void CarregarListas() {
            Execute(d => {
                Model.Listas = d.CarregarListas();
            });
            
        }

        public void CarregarProximaLista() {
            Execute(d => {
                Model.ProximaLista = d.CarregarProximaLista();
            });
        }

        public int ContagemCandidatos() {
            int contagemCandidatos = 0;
            Execute(d => {
                contagemCandidatos = d.ContagemCandidatos();
            });
            return contagemCandidatos;
        }

        public void AtualizarListas() {
            Execute(d => {
                d.AtualizarListas(Model.Listas);
                AtualizarStatusSorteio(d, Status.SORTEIO);
            });
        }

        public void CriarListasSorteioDeFaixas(string arquivoImportacao, string faixa, Action<string> updateStatus, Action<int> updateProgress, int listaAtual, int totalListas, int incremento)
        {
            if (arquivoImportacao != null)
            {
                Execute(d =>
                {
                    using (Stream stream = File.OpenRead(arquivoImportacao))
                    {
                        using (IExcelDataReader excelReader = CreateExcelReader(arquivoImportacao, stream))
                        {
                            d.CopiarCandidatosArquivo(faixa, excelReader, updateStatus, updateProgress);
                        }
                    }
                });
            }
            Execute(d => {
                d.CriarListasSorteioPorFaixa(faixa, updateStatus, updateProgress, listaAtual, totalListas, incremento);
                AtualizarStatusSorteio(d, Status.QUANTIDADES);
            });
        }

        private IExcelDataReader CreateExcelReader(string arquivoImportacao, Stream stream) {
            return (arquivoImportacao.ToLower().EndsWith(".xlsx") || arquivoImportacao.ToLower().EndsWith(".xls")) ?
                ExcelReaderFactory.CreateOpenXmlReader(stream) : ExcelReaderFactory.CreateBinaryReader(stream);
        }

        public void SortearProximaLista(Action<string> updateStatus, Action<int> updateProgress, Action<string> logText, int? sementePersonalizada = null) {
            Execute(d => {
                d.SortearProximaLista(updateStatus, updateProgress, logText, sementePersonalizada);
                if (Model.StatusSorteio == Status.SORTEIO) {
                    AtualizarStatusSorteio(d, Status.SORTEIO_INICIADO);
                }
                if (d.CarregarProximaLista() == null) {
                    AtualizarStatusSorteio(d, Status.FINALIZADO);
                }
            });
        }

        public string DiretorioExportacaoCSV => Database.DiretorioExportacaoCSV;
        public bool DiretorioExportacaoCSVExistente => Directory.Exists(Database.DiretorioExportacaoCSV);

        public void ExportarListas(Action<string> updateStatus) {
            Execute(d => {
                d.ExportarListas(updateStatus);
            });
        }

        public void SalvarLista(Lista lista, string caminhoArquivo) {
            ListaPub listaPublicacao = null;
            Execute(d => { listaPublicacao = d.CarregarListaPublicacao(lista.IdLista); });
            PdfFileWriter.WriteToPdf(caminhoArquivo, Model, listaPublicacao);
            System.Diagnostics.Process.Start(caminhoArquivo);
        }

        public void SalvarSorteados(string caminhoArquivo)
        {
            ListaPub listaPublicacao = null;
            Execute(d => { listaPublicacao = d.CarregarListaSorteados(); });
            PdfFileWriter.WriteSorteadosToPdf(caminhoArquivo, Model, listaPublicacao);
            System.Diagnostics.Process.Start(caminhoArquivo);
        }
    }
}
