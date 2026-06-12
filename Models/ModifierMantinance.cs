using Microsoft.AspNetCore.Mvc;

namespace MSaleWebServer.Models
{
    public class ModifierNameDto
    {
        public string ModifierName { get; set; }
    }

    public class ModifierItemDto
    {
        public string ModifierCategory { get; set; }
        public int Pos { get; set; }
    }

    public class SaveModifierItemsDto
    {
        public string ModifierName { get; set; }
        public List<ModifierItemDto> Items { get; set; }
    }
}
