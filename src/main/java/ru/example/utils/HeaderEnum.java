package ru.example.utils;

public enum HeaderEnum {
  DATE("Дата"),
  HOST_ID("Идентификатор хоста"),
  DISC_MODEL("Модель диска"),
  DISC_SERIAL_NUMBER("Серийный номер диска"),
  DISC_TOTAL_SIZE("Полный объем диска"),
  HOURS_IN_OPERATION("Часов в работе"),
  RESTART_CYCLE("Циклов перезапуска"),
  MB_CHECK("МБайт записано"),
  MB_READ("МБайт прочитано"),
  SIZE_IN_BYTES("Размер в байтах"),
  TEMPERATURE("Температура");

  private final String key;

  HeaderEnum(String key) {
    this.key = key;
  }

  public String getKey() {
    return key;
  }
}
