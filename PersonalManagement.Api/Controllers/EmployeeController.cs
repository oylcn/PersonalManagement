using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using MediatR;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using PersonalManagement.Application.Features.Employees.Commands.CreateEmployee;
using PersonalManagement.Application.Features.Employees.Commands.DeleteEmployee;
using PersonalManagement.Application.Features.Employees.Commands.UpdateEmployee;
using PersonalManagement.Application.Features.Employees.Queries.GetEmployeeDetail;
using PersonalManagement.Application.Features.Employees.Queries.GetEmployeesList;

namespace PersonalManagement.Api.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class EmployeeController : ControllerBase
    {
        private readonly IMediator _mediator;
        public EmployeeController(IMediator mediator)
        {
            _mediator = mediator;
        }

        //[Authorize]
        [HttpGet("all", Name = "GetEmployees")]
        [ProducesResponseType(StatusCodes.Status200OK)]
        [ProducesDefaultResponseType]
        public async Task<ActionResult<List<EmployeeListVm>>> GetEmployees()
        {
            var dtos = await _mediator.Send(new GetEmployeesListQuery());
            return Ok(dtos);
        }

        [HttpGet("{id}", Name = "GetEmployeeById")]
        public async Task<ActionResult<EmployeeDetailVm>> GetEmployeeById(int id)
        {
            var getEmployeeDetailQuery = new GetEmployeeDetailQuery() { EmployeeId = id };
            return Ok(await _mediator.Send(getEmployeeDetailQuery));
        }

        [HttpPost(Name = "AddEmployee")]
        public async Task<ActionResult<CreateEmployeeCommandResponse>> Create([FromBody] CreateEmployeeCommand createEmployeeCommand)
        {
            var response = await _mediator.Send(createEmployeeCommand);
            return Ok(response);
        }

        [HttpPut(Name = "UpdateEmployee")]
        [ProducesResponseType(StatusCodes.Status204NoContent)]
        [ProducesResponseType(StatusCodes.Status404NotFound)]
        [ProducesDefaultResponseType]
        public async Task<ActionResult> Update([FromBody] UpdateEmployeeCommand updateEmployeeCommand)
        {
            await _mediator.Send(updateEmployeeCommand);
            return NoContent();
        }

        [HttpDelete("{id}", Name = "DeleteEmployee")]
        [ProducesResponseType(StatusCodes.Status204NoContent)]
        [ProducesResponseType(StatusCodes.Status404NotFound)]
        [ProducesDefaultResponseType]
        public async Task<ActionResult> Delete(int id)
        {
            var deleteEmployeeCommand = new DeleteEmployeeCommand() { EmployeeId = id };
            await _mediator.Send(deleteEmployeeCommand);
            return NoContent();
        }
    }
}
