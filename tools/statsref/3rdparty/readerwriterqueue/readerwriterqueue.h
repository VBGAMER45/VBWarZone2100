// statsref's stand-in for moodycamel's readerwriterqueue (the submodule is not checked out): a plain FIFO. statsref has
// no threads, so netreplay.cpp's writer drains it where NETreplaySaveStop joins the thread (see run_tests.py).
#pragma once
#include <cstddef>
#include <deque>
#include <utility>

namespace moodycamel
{
template <typename T>
class BlockingReaderWriterQueue
{
public:
	explicit BlockingReaderWriterQueue(size_t = 15)
	{
	}
	bool enqueue(T &&item)
	{
		q.push_back(std::move(item));
		return true;
	}
	bool enqueue(T const &item)
	{
		q.push_back(item);
		return true;
	}
	void wait_dequeue(T &item)
	{
		item = std::move(q.front());
		q.pop_front();
	}
private:
	std::deque<T> q;
};
}
