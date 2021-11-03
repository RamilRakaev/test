using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using RecruitingStaff.Domain.Model.CandidateQuestionnaire;
using RecruitingStaff.Infrastructure.Repositories;
using Microsoft.EntityFrameworkCore;
using RecruitingStaff.Domain.Model;
using RecruitingStaff.Domain.Interfaces;
using RecruitingStaff.Infrastructure.Repositories.SubRepositories;

namespace test
{

    public class SaveParseQuestionnare
    {
        private readonly IRepository<Candidate> _candidateRepository;
        private readonly IRepository<Vacancy> _vacancyRepository;
        private readonly IRepository<CandidateVacancy> _candidateVacancyRepository;
        private readonly IRepository<RecruitingStaffWebAppFile> _fileRepository;
        private readonly IRepository<Questionnaire> _questionnaireRepository;
        private readonly IRepository<QuestionCategory> _questionCategoryRepository;
        private readonly IRepository<Question> _questionRepository;
        private readonly IRepository<Answer> _answerRepository;

        private RecruitingStaffWebAppFile _file;
        private Vacancy currentVacancy;
        private Candidate currentCandidate;
        private Questionnaire currentQuestionnaire;
        private QuestionCategory currentCategory;

        public SaveParseQuestionnare(
            IRepository<Vacancy> vacancyRepository,
            IRepository<Candidate> candidateRepository,
            IRepository<CandidateVacancy> candidateVacancyRepository,
            IRepository<RecruitingStaffWebAppFile> fileRepository,
            IRepository<Questionnaire> questionnaireRepository,
            IRepository<QuestionCategory> questionCategoryRepository,
            IRepository<Question> questionRepository,
            IRepository<Answer> answerRepository)
        {
            _vacancyRepository = vacancyRepository;
            _candidateRepository = candidateRepository;
            _candidateVacancyRepository = candidateVacancyRepository;
            _fileRepository = fileRepository;
            _questionnaireRepository = questionnaireRepository;
            _questionCategoryRepository = questionCategoryRepository;
            _questionRepository = questionRepository;
            _answerRepository = answerRepository;
        }

        public async Task Parse(string document)
        {
            _file = new RecruitingStaffWebAppFile() { Source = document, FileType = FileType.Questionare };
            await _fileRepository.AddAsync(_file);
            await _fileRepository.SaveAsync();
            using (var wordDoc = WordprocessingDocument.Open(_file.Source, false))
            {
                var body = wordDoc.MainDocumentPart.Document.Body;
                currentQuestionnaire = new Questionnaire()
                { Name = body.ChildElements.Where(e => e.LocalName == "p").FirstOrDefault().InnerText };

                foreach (var element in body.ChildElements.Where(e => e.LocalName == "tbl"))
                {
                    await ParseCandidateWithVacancy(element);
                    foreach (var row in element.ChildElements.Reverse())
                    {
                        foreach (var cell in row.ChildElements)
                        {
                            var table = cell.FirstOrDefault(c => c.LocalName == "tbl");
                            if (table != null)
                            {
                                await ParseQuestionnaire(table);
                            }
                        }
                    }
                }
            }
        }

        private async Task ParseCandidateWithVacancy(OpenXmlElement table)
        {
            currentCandidate = new Candidate();
            var rows = table.ChildElements.Where(e => e.LocalName == "tr");
            var name = rows.ElementAt(0).InnerText;
            var vacancyName = name[(name.IndexOf(':') + 2)..];

            currentVacancy = _vacancyRepository.GetAll().Where(v => v.Name == vacancyName).FirstOrDefault();
            currentCandidate.FullName = ExtractCellTextFromRow(rows, 1, 1);
            try
            {
                var dateStr = ExtractCellTextFromRow(rows, 2, 1);
                currentCandidate.DateOfBirth = dateStr != string.Empty ? Convert.ToDateTime(dateStr) : new DateTime();
            }
            catch
            {
                currentCandidate.DateOfBirth = new DateTime();
            }

            if (currentVacancy == null)
            {
                currentVacancy = new Vacancy() { Name = vacancyName };
                await _vacancyRepository.AddAsync(currentVacancy);
                await _vacancyRepository.SaveAsync();
            }

            currentCandidate.Address = ExtractCellTextFromRow(rows, 2, 2);
            currentCandidate.Address = currentCandidate.Address[(currentCandidate.Address.IndexOf(':') + 2)..];
            currentCandidate.TelephoneNumber = ExtractCellTextFromRow(rows, 3, 1);
            currentCandidate.MaritalStatus = ExtractCellTextFromRow(rows, 4, 1);

            await _candidateRepository.AddAsync(currentCandidate);
            await _candidateRepository.SaveAsync();

            var candidateVacancy = new CandidateVacancy()
            {
                CandidateId = currentCandidate.Id,
                VacancyId = currentVacancy.Id
            };
            await _candidateVacancyRepository.AddAsync(candidateVacancy);
            await _candidateVacancyRepository.SaveAsync();
        }

        private static string ExtractCellTextFromRow(IEnumerable<OpenXmlElement> rows, int rowInd, int cellInd)
        {
            return rows
                .ElementAt(rowInd)
                .ChildElements
                .Where(e => e.LocalName == "tc")
                .ElementAt(cellInd).InnerText;
        }

        private async Task ParseQuestionnaire(OpenXmlElement table)
        {
            currentQuestionnaire.CandidateId = currentCandidate.Id;
            currentQuestionnaire.VacancyId = currentVacancy.Id;
            currentQuestionnaire.DocumentFileId = _file.Id;

            await _questionnaireRepository.AddAsync(currentQuestionnaire);
            await _questionnaireRepository.SaveAsync();

            foreach (var child in table.ChildElements.Where(e => e.LocalName == "tr").Skip(1))
            {
                if (child.ChildElements.Count == 3)
                {
                    currentCategory = new QuestionCategory()
                    {
                        Name = child.ChildElements.ElementAt(2).InnerText,
                        QuestionnaireId = currentQuestionnaire.Id
                    };
                    await _questionCategoryRepository.AddAsync(currentCategory);
                    await _questionCategoryRepository.SaveAsync();
                }
                if (child.ChildElements.Count == 5)
                {
                    await ParseQuestion(child);
                }
            }
        }

        private async Task ParseQuestion(OpenXmlElement child)
        {
            var question = new Question
            {
                QuestionCategoryId = currentCategory.Id,
                Name = child.ChildElements[2].InnerText
            };
            await _questionRepository.AddAsync(question);
            await _questionRepository.SaveAsync();
            if (child.ChildElements[4].InnerText != string.Empty)
            {
                var answer = new Answer
                {
                    CandidateId = currentCandidate.Id,
                    QuestionId = question.Id,
                    Comment = child.ChildElements[4].InnerText
                };
                try
                {
                    answer.Estimation = child.ChildElements[3].InnerText == string.Empty ?
                        (byte)0 : Convert.ToByte(child.ChildElements[3].InnerText);
                }
                catch
                { }
                await _answerRepository.AddAsync(answer);
                await _answerRepository.SaveAsync();
            }
        }
    }

    public class Program
    {
        static async Task Main(string[] args)
        {
            var builder = new DbContextOptionsBuilder<DataContext>();
            var options = builder.UseNpgsql("Host=localhost;Port=5432;Database=recrutingstaffdb;Username=postgres;Password=rubaka").Options;
            var db = new DataContext(options);

            ResetDb(db);
            var dbparser = new SaveParseQuestionnare(
                new VacancyRepository(db),
                new CandidateRepository(db),
                new CandidateVacancyRepository(db),
                new RecruitingStaffWebAppFileRepository(db),
                new QuestionnaireRepository(db),
                new QuestionCategoryRepository(db),
                new QuestionRepository(db),
                new AnswerRepository(db));
            await dbparser.Parse(@"C:\Users\Public\Downloads\Анкета IT (4).docx");
            Console.WriteLine("end!");
            Console.ReadLine();
        }

        public static void ResetDb(DataContext db)
        {
            ResetDb<Answer>(db);
            ResetDb<Question>(db);
            ResetDb<QuestionCategory>(db);
            ResetDb<Questionnaire>(db);
            ResetDb<CandidateVacancy>(db);
            ResetDb<Vacancy>(db);
            ResetDb<Candidate>(db);
            ResetDb<RecruitingStaffWebAppFile>(db);
        }

        public static void ResetDb<Entity>(DataContext db) where Entity : BaseEntity
        {
            var entities = db.Set<Entity>().ToArray();
            var f = entities.Count();
            for (int i = 0; i < entities.Count(); i++)
            {
                db.Set<Entity>().Remove(entities.ElementAt(i));
                db.SaveChanges();
            }
        }
    }

    public class ParseQuestionnare
    {
        int index = 1;
        private readonly IRepository<Candidate> _candidateRepository;
        private readonly IRepository<Vacancy> _vacancyRepository;
        private readonly IRepository<CandidateVacancy> _candidateVacancyRepository;
        private readonly IRepository<Questionnaire> _questionnaireRepository;
        private readonly IRepository<QuestionCategory> _questionCategoryRepository;
        private readonly IRepository<Question> _questionRepository;
        private readonly IRepository<Answer> _answerRepository;

        private RecruitingStaffWebAppFile _file;
        private Vacancy currentVacancy;
        private Candidate currentCandidate;
        private Questionnaire currentQuestionnaire;

        public ParseQuestionnare(
            IRepository<Vacancy> vacancyRepository,
            IRepository<Candidate> candidateRepository,
            IRepository<CandidateVacancy> candidateVacancyRepository,
            IRepository<Questionnaire> questionnaireRepository,
            IRepository<QuestionCategory> questionCategoryRepository,
            IRepository<Question> questionRepository,
            IRepository<Answer> answerRepository)
        {
            _vacancyRepository = vacancyRepository;
            _candidateRepository = candidateRepository;
            _candidateVacancyRepository = candidateVacancyRepository;
            _questionnaireRepository = questionnaireRepository;
            _questionCategoryRepository = questionCategoryRepository;
            _questionRepository = questionRepository;
            _answerRepository = answerRepository;
        }

        public ParseQuestionnare()
        {

        }

        public async Task Parse(string document)
        {
            _file = new RecruitingStaffWebAppFile() { Source = document, FileType = FileType.Questionare, Id = index++ };
            using (var wordDoc = WordprocessingDocument.Open(_file.Source, false))
            {
                var body = wordDoc.MainDocumentPart.Document.Body;
                currentQuestionnaire = new Questionnaire()
                { Name = body.ChildElements.Where(e => e.LocalName == "p").FirstOrDefault().InnerText };

                foreach (var element in body.ChildElements.Where(e => e.LocalName == "tbl"))
                {
                    ParseCandidateWithVacancy(element);
                    foreach (var row in element.ChildElements)
                    {
                        foreach (var cell in row.ChildElements)
                        {
                            var table = cell.FirstOrDefault(c => c.LocalName == "tbl");
                            if (table != null)
                            {
                                ParseQuestions(table);
                            }
                        }
                    }
                }
            }
            //add questionnaire
        }

        private void ParseCandidateWithVacancy(OpenXmlElement table)
        {
            currentCandidate = new Candidate() { Id = index++ };
            var rows = table.ChildElements.Where(e => e.LocalName == "tr");
            var name = rows.ElementAt(0).InnerText;
            var vacancyName = name[(name.IndexOf(':') + 2)..];

            //vacancy first or default

            currentCandidate.FullName = ExtractCellTextFromRow(rows, 1, 1);
            try
            {
                var dateStr = ExtractCellTextFromRow(rows, 2, 1);
                currentCandidate.DateOfBirth = dateStr != string.Empty ? Convert.ToDateTime(dateStr) : new DateTime();
            }
            catch
            {
                currentCandidate.DateOfBirth = new DateTime();
            }
            //add candidate
            if (currentVacancy == null)
            {
                currentVacancy = new Vacancy() { Name = vacancyName, Id = index++ };
                var candidateVacancy = new CandidateVacancy()
                {
                    CandidateId = currentCandidate.Id,
                    VacancyId = currentVacancy.Id
                };
                //add CandidateVacancy
            }
            currentCandidate.Address = ExtractCellTextFromRow(rows, 2, 2);
            currentCandidate.Address = currentCandidate.Address[(currentCandidate.Address.IndexOf(':') + 2)..];
            currentCandidate.TelephoneNumber = ExtractCellTextFromRow(rows, 3, 1);
            currentCandidate.MaritalStatus = ExtractCellTextFromRow(rows, 4, 1);
        }

        private static string ExtractCellTextFromRow(IEnumerable<OpenXmlElement> rows, int rowInd, int cellInd)
        {
            return rows
                .ElementAt(rowInd)
                .ChildElements
                .Where(e => e.LocalName == "tc")
                .ElementAt(cellInd).InnerText;
        }

        private void ParseQuestions(OpenXmlElement table)
        {
            currentQuestionnaire.Id = index++;
            currentQuestionnaire.CandidateId = currentCandidate.Id;
            currentQuestionnaire.VacancyId = currentVacancy.Id;
            currentQuestionnaire.DocumentFileId = _file.Id;

            var questionCategories = new List<QuestionCategory>();
            var questions = new List<Question>();
            var answers = new List<Answer>();
            var currentCategory = new QuestionCategory();

            foreach (var child in table.ChildElements.Where(e => e.LocalName == "tr").Skip(1))
            {
                if (child.ChildElements.Count == 3)
                {
                    currentCategory = new QuestionCategory()
                    {
                        Name = child.ChildElements.ElementAt(2).InnerText,
                        Id = index++,
                        QuestionnaireId = currentQuestionnaire.Id
                    };
                    questionCategories.Add(currentCategory);
                }
                if (child.ChildElements.Count == 5)
                {
                    var question = new Question
                    {
                        Id = index++,
                        QuestionCategoryId = currentCategory.Id,
                        Name = child.ChildElements[2].InnerText
                    };
                    questions.Add(question);
                    question = questions.First(q => q == question);
                    if (child.ChildElements[4].InnerText != string.Empty)
                    {
                        var answer = new Answer
                        {
                            QuestionId = question.Id,
                            CandidateId = currentCandidate.Id,
                            Estimation =
                            child.ChildElements[3].InnerText == string.Empty ?
                            (byte)0 : Convert.ToByte(child.ChildElements[3].InnerText),
                            Comment = child.ChildElements[4].InnerText
                        };
                        answers.Add(answer);
                    }
                }
            }
        }

        private void ParseQuestion(OpenXmlElement child, QuestionCategory currentCategory)
        {
            var question = new Question
            {
                QuestionCategoryId = currentCategory.Id,
                Name = child.ChildElements[2].InnerText
            };
            if (child.ChildElements[4].InnerText != string.Empty)
            {
                var answer = new Answer
                {
                    QuestionId = question.Id,
                    Estimation =
                    child.ChildElements[3].InnerText == string.Empty ?
                    (byte)0 : Convert.ToByte(child.ChildElements[3].InnerText),
                    Comment = child.ChildElements[4].InnerText
                };
                _answerRepository.AddAsync(answer);
            }
        }
    }
}

