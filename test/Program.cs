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

namespace test
{
    public class ParseQuestionnare
    {
        private readonly IRepository<Candidate> _candidateRepository;
        private readonly IRepository<Questionnaire> _questionnaireRepository;
        private readonly IRepository<Vacancy> _vacancyRepository;

        public ParseQuestionnare(
            IRepository<Candidate> candidateRepository,
            IRepository<Questionnaire> questionnaireRepository,
            IRepository<Vacancy> vacancyRepository)
        {
            _candidateRepository = candidateRepository;
            _questionnaireRepository = questionnaireRepository;
            _vacancyRepository = vacancyRepository;
        }

        public async Task Parse(string document)
        {
            var questionnare = new Questionnaire();
            using (var wordDoc = WordprocessingDocument.Open(document, false))
            {
                MainDocumentPart mainPart = wordDoc.MainDocumentPart;
                foreach (var element in mainPart.Document.Body.ChildElements.Where(e => e.LocalName == "tbl"))
                {
                    await ParseCandidate(element);
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
            await _questionnaireRepository.AddAsync(questionnare);
        }

        private async Task ParseCandidateWithVacancy(OpenXmlElement table)
        {
            var candidate = new Candidate();
            var rows = table.ChildElements.Where(e => e.LocalName == "tr");
            var name = rows.ElementAt(0).InnerText;
            var vacancyName = name[(name.IndexOf(':') + 1)..];
            var vacancy = _vacancyRepository
                .GetAllAsNoTracking()
                .FirstOrDefault(v => v.Name == vacancyName);
            if (vacancy == null)
            {
                vacancy = new Vacancy() { Name = vacancyName };
                await _vacancyRepository.AddAsync(vacancy);
                await _vacancyRepository.SaveAsync();
            }

            candidate.FullName = ExtractCellTextFromRow(rows, 1, 1);
            var dateStr = ExtractCellTextFromRow(rows, 2, 1);
            candidate.DateOfBirth = dateStr != string.Empty ? Convert.ToDateTime(dateStr) : new DateTime();
            await _candidateRepository.AddAsync(candidate);
        }
        private string ExtractCellTextFromTable(OpenXmlElement table, int row, int cell)
        {
            return table.ChildElements
                .Where(e => e.LocalName == "tr")
                .ElementAt(row)
                .Where(e => e.LocalName == "tc")
                .ElementAt(cell).InnerText;
        }

        private string ExtractCellTextFromRow(IEnumerable<OpenXmlElement> rows, int rowInd, int cellInd)
        {
            return rows
                .ElementAt(rowInd)
                .ChildElements
                .Where(e => e.LocalName == "tc")
                .ElementAt(cellInd).InnerText;
        }

        private void ParseQuestions(OpenXmlElement table)
        {
            var questionCategories = new List<QuestionCategory>();
            var questions = new List<Question>();
            var answers = new List<Answer>();
            var currentCategory = new QuestionCategory();
            foreach (var child in table.ChildElements.Where(e => e.LocalName == "tr").Skip(1))
            {
                if (child.ChildElements.Count == 3)
                {
                    currentCategory = new QuestionCategory();
                    questionCategories.Add(currentCategory);
                    currentCategory = questionCategories.First(c => c == currentCategory);
                }
                if (child.ChildElements.Count == 5)
                {
                    var question = new Question
                    {
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
    }

    public class Program
    {
        static void Main(string[] args)
        {
            var parser = new ParseQuestionnare();
            parser.Parse(@"C:\Users\Public\Downloads\Анкета IT (4).docx");
            Console.ReadLine();
        }


    }
}

