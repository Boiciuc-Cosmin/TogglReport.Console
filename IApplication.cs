using System.Threading.Tasks;
using TogglReport.ConsoleApp.Dtos;

namespace TogglReport.ConsoleApp {
    public interface IApplication {
        Task RunAsync(ArgumentOptionsModel argumentOptions);
    }
}