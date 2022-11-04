package cn.creekmoon.excelUtils.hutool589.core.thread.lock;

import java.util.concurrent.locks.ReentrantReadWriteLock;
import java.util.concurrent.locks.StampedLock;

/**
 * 锁相关工具
 *
 * @author looly
 * @since 5.2.5
 */
public class LockUtil {

	private static final cn.creekmoon.excelUtils.hutool589.core.thread.lock.NoLock NO_LOCK = new cn.creekmoon.excelUtils.hutool589.core.thread.lock.NoLock();

	/**
	 * 创建{@link StampedLock}锁
	 *
	 * @return {@link StampedLock}锁
	 */
	public static StampedLock createStampLock() {
		return new StampedLock();
	}

	/**
	 * 创建{@link ReentrantReadWriteLock}锁
	 *
	 * @param fair 是否公平锁
	 * @return {@link ReentrantReadWriteLock}锁
	 */
	public static ReentrantReadWriteLock createReadWriteLock(boolean fair) {
		return new ReentrantReadWriteLock(fair);
	}

	/**
	 * 获取单例的无锁对象
	 *
	 * @return {@link cn.creekmoon.excelUtils.hutool589.core.thread.lock.NoLock}
	 */
	public static NoLock getNoLock() {
		return NO_LOCK;
	}
}
