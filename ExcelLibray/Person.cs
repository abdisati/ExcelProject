namespace ExcelLibray;

public class Person
{
  public string? Name { get; set; }
    public int Age { get; set; }
    public string? City { get; set; }

    public override string ToString()
    {
        return $"Name: {Name}, Age: {Age}, City: {City}";
    }
}
