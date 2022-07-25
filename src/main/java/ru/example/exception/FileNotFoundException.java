package ru.example.exception;

public class FileNotFoundException extends RuntimeException {

  public FileNotFoundException(String message) {
    super(String.format("конечная директория не найдена или пуста: %s", message));
  }

}
