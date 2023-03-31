#!/bin/bash
echo "正在切换目录..."
cd $(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)/half
echo "正在执行文件..."
echo "Mac平台上可能需要十多秒甚至更长时间加载webdriver，请耐心等待"
./half
read -n 1 -s -r -p "按任意键退出"
