const robot = require('robotjs'); // 确保安装了 robotjs

process.argv.slice(2).forEach((inputText) => {
    for (let index = 0; index < parseInt(inputText); index++) {

        robot.moveMouse(400, 120);
        // console.log(robot.getPixelColor(400, 120))
        // const screenSize = robot.getScreenSize();
        // console.log(screenSize);
        // robot.setMouseDelay(1000);
        // robot.mouseClick();
        // robot.moveMouseSmooth(xxx, yyy);

        // 重复其他操作
        process.send("Completed");
    }
});
