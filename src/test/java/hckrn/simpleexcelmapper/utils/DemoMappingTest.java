package hckrn.simpleexcelmapper.utils;

import hckrn.simpleexcelmapper.utils.dto.Demo;
import lombok.SneakyThrows;
import org.apache.poi.ss.usermodel.Workbook;
import org.assertj.core.api.Assertions;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.security.SecureRandom;
import java.time.LocalDate;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.LinkedList;
import java.util.List;

class DemoMappingTest {

    private final SecureRandom random = new SecureRandom();


    @Test
    @SneakyThrows
    void shouldCreateDemoReport() {
        var excelMapper = new ExcelMapperImpl();

        var fileName = "demo-out-" + LocalTime.now().format(DateTimeFormatter.ofPattern("hh.mm.ss")) + ".xlsx";
        List<Demo> demos = generateDemos();

        try (Workbook workbook = excelMapper.createWorkbookFromObject(demos);
             var fileOutputStream = new FileOutputStream(fileName)) {
            workbook.write(fileOutputStream);
        }
        Assertions.assertThat(new File(fileName)).exists();
        Files.delete(Path.of(fileName));
    }

    private List<Demo> generateDemos() {
        LinkedList<Demo> demos = new LinkedList<>();
        for (int i = 1; i <= 100; i++) {
            demos.add(new Demo()
                    .setRef(i)
                    .setDate(LocalDate.now().minusMonths(random.nextInt(0, 13)))
                    .setHome(getRandomTeam())
                    .setAway(getRandomTeam())
                    .setLeague(getRandomLeague())
                    .setFtOver05(random.nextDouble(0, 1))
                    .setFtOver15(random.nextDouble(0, 1))
                    .setFtOver25(random.nextDouble(0, 1))
                    .setFtOver35(random.nextDouble(0, 1))
                    .setFtBtt(random.nextDouble(0, 1))
                    .setHomeGoalScores(random.nextDouble(1, 2))
                    .setHomeGoalsCc(random.nextDouble(1, 2))
                    .setAwayGoalScores(random.nextDouble(1, 2))
                    .setAwayGoalsCc(random.nextDouble(1, 2))
                    .setHomeGoalsScored(random.nextDouble(0, 1)));
        }
        return demos;
    }

    private String getRandomTeam() {
        List<String> teams = List.of("Harbor City Seagulls",
                "Redwood United",
                "Mountain Ridge Rovers",
                "Coastal Waves FC",
                "Golden Plains Wanderers",
                "Phoenix Blaze FC",
                "Uptown Eagles",
                "Starlight Strikers",
                "Lunar Valley United",
                "Riverside Rhinos",
                "Kingsport Knights FC",
                "Thunderpeak Thunderbolts",
                "Ironhill Invincibles",
                "Forest Glen Foxes",
                "Suncrest Sapphires FC",
                "Polar Point Panthers",
                "Eastside Emperors",
                "Westwind Warriors",
                "Serenity Sands FC",
                "Twilight Titans"
        );

        return teams.get(random.nextInt(0, teams.size()));
    }

    private String getRandomLeague() {
        List<String> leagues = List.of(
                "Northern Star Premier League",
                "Golden Crest Championship",
                "Eclipse Elite Division",
                "Polaris Pro Series",
                "Midland Masters Circuit",
                "Sunrise Super League",
                "Astral First Division",
                "Equator Excellence League",
                "Twilight Top Tier",
                "Zenith Zone Championship"
        );

        return leagues.get(random.nextInt(0, leagues.size()));
    }
}
