using System;
using System.Collections.Generic;
using System.Threading.Tasks;
namespace BlazorCRUDsf.Data
{
    // Each item below provides an interface to a method in ColourServices.cs
    public interface IColourService
    {
        Task<bool> ColourInsert(Colour colour);
        Task<IEnumerable<Colour>> ColourList();
        Task<Colour> Colour_GetOne(string Code);
        Task<bool> ColourUpdate(Colour colour);
        Task<bool> ColourDelete(string Code);

    }
}
